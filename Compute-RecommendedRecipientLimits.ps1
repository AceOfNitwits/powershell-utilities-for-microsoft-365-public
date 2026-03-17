# Utility: page through Get-MessageTraceV2 results
#
# Exchange Online's Get-MessageTraceV2 API returns results in pages and supports a
# "StartingRecipientAddress" bookmark. This helper abstracts the paging logic so
# callers can request a full date range and receive a complete collection of rows.
function Get-MessageTraceV2All {
    param(
        [Parameter(Mandatory = $true)][datetime]$StartDate,
        [Parameter(Mandatory = $true)][datetime]$EndDate,
        [ValidateRange(1, 5000)][int]$ResultSize = 5000,
        [ValidateRange(1, 5000)][int]$MaxRounds = 2000,
        [ValidateRange(0, 60000)][int]$SleepMsBetweenRounds = 300
    )

    $all = New-Object System.Collections.Generic.List[object]
    $currentEnd = $EndDate
    $startingRecipient = $null

    for ($round = 1; $round -le $MaxRounds; $round++) {
        $params = @{
            StartDate  = $StartDate
            EndDate    = $currentEnd
            ResultSize = $ResultSize
        }
        if ($startingRecipient) {
            $params.StartingRecipientAddress = $startingRecipient
        }

        $batch = Get-MessageTraceV2 @params
        if (-not $batch -or $batch.Count -eq 0) { break }

        foreach ($row in $batch) { $all.Add($row) }

        if ($batch.Count -lt $ResultSize) { break }

        $last = $batch[-1]
        $lastReceived = [datetime]$last.Received
        $lastRecipient = [string]$last.RecipientAddress

        $currentEnd = if ($lastReceived -ge $currentEnd) { $currentEnd.AddSeconds(-1) } else { $lastReceived }
        $startingRecipient = $lastRecipient

        if ($SleepMsBetweenRounds -gt 0) { Start-Sleep -Milliseconds $SleepMsBetweenRounds }
    }

    return $all
}

# Computes recommended recipient limits based on observed message trace volumes.
#
# This function:
#   1) Retrieves message trace data for the specified lookback window.
#   2) Counts recipients per sender per hour (split internal vs external) and per day.
#   3) Computes the historical peak for each sender and then applies a multiplier
#      and minimums to arrive at recommended recipient limits.
function Compute-MaxRecipientLimitsV2 {
    param(
        [int]$DaysBack = 30,
        [ValidateRange(1, 10)][int]$ChunkDays = 5,
        [ValidateRange(1, 5000)][int]$ResultSize = 5000,

        [ValidateRange(1.0, 5.0)][double]$Multiplier = 1.25,

        [int]$MinExternalHourly = 25,
        [int]$MinInternalHourly = 25,
        [int]$MinDaily          = 100,

        [string[]]$ValidSenderDomains = @()
    )

    # Domains considered "internal" for the purposes of counting internal vs external recipients.
    $acceptedDomains = (Get-AcceptedDomain).DomainName | ForEach-Object { $_.ToLowerInvariant() }

    # The list of sender domains we will include in the report.
    # If the caller provides a list via -ValidSenderDomains, use that; otherwise default to all accepted domains.
    $validDomains = if ($ValidSenderDomains -and $ValidSenderDomains.Count -gt 0) {
        $ValidSenderDomains | ForEach-Object { $_.ToLowerInvariant() }
    } else {
        $acceptedDomains
    }

    $start = (Get-Date).AddDays(-1 * $DaysBack)
    $end   = Get-Date

    $extHourCounts = @{}   # key senderAddress|hour
    $intHourCounts = @{}   # key senderAddress|hour
    $dayTotalCounts = @{}  # key senderAddress|day (internal + external)

    # Iterate over the requested date range in manageable "chunks".
    # This keeps each Get-MessageTraceV2 call from requesting too much data at once.
    $cursor = $start
    while ($cursor -lt $end) {
        $chunkStart = $cursor
        $chunkEnd = $cursor.AddDays($ChunkDays)
        if ($chunkEnd -gt $end) { $chunkEnd = $end }

        $rows = Get-MessageTraceV2All -StartDate $chunkStart -EndDate $chunkEnd -ResultSize $ResultSize

        foreach ($r in $rows) {
            if (-not $r.SenderAddress -or -not $r.RecipientAddress -or -not $r.Received) { continue }

            $senderAddress = $r.SenderAddress.ToLowerInvariant()
            $rcpt   = $r.RecipientAddress.ToLowerInvariant()

            $senderAt = $senderAddress.LastIndexOf("@")
            if ($senderAt -lt 0) { continue }
            $senderDomain = $senderAddress.Substring($senderAt + 1)
            if ($validDomains -notcontains $senderDomain) { continue }

            $rcptAt = $rcpt.LastIndexOf("@")
            if ($rcptAt -lt 0) { continue }
            $rcptDomain = $rcpt.Substring($rcptAt + 1)

            # Decide whether the recipient is internal (accepted domain) vs external.
            $isInternalRecipient = ($acceptedDomains -contains $rcptDomain)

            # Group messages into hourly and daily "buckets" for counting.
            # Hourly counts are split between internal and external recipients.
            $dtLocal = ([datetime]$r.Received).ToLocalTime()
            $hourBucket = $dtLocal.ToString("yyyy-MM-dd HH")
            $dayBucket  = $dtLocal.ToString("yyyy-MM-dd")

            $hk = "$senderAddress|$hourBucket"
            $dk = "$senderAddress|$dayBucket"

            if ($isInternalRecipient) {
                if ($intHourCounts.ContainsKey($hk)) { $intHourCounts[$hk]++ } else { $intHourCounts[$hk] = 1 }
            } else {
                if ($extHourCounts.ContainsKey($hk)) { $extHourCounts[$hk]++ } else { $extHourCounts[$hk] = 1 }
            }

            # Total daily recipients (internal + external) is used for the daily limit recommendation.
            if ($dayTotalCounts.ContainsKey($dk)) { $dayTotalCounts[$dk]++ } else { $dayTotalCounts[$dk] = 1 }
        }

        $cursor = $chunkEnd
    }

    # Build a set of unique senders encountered in the message trace.
    # Each count dictionary uses keys of the form <sender>|<bucket>, so we split to get the sender.
    $senders = New-Object System.Collections.Generic.HashSet[string]
    foreach ($k in $extHourCounts.Keys) { [void]$senders.Add($k.Split("|")[0]) }
    foreach ($k in $intHourCounts.Keys) { [void]$senders.Add($k.Split("|")[0]) }
    foreach ($k in $dayTotalCounts.Keys) { [void]$senders.Add($k.Split("|")[0]) }

    $rowsOut = foreach ($s in ($senders | Sort-Object)) {
        $maxExtH = 0
        $maxIntH = 0
        $maxDay  = 0

        foreach ($k in $extHourCounts.Keys) {
            if ($k.StartsWith("$s|")) { if ($extHourCounts[$k] -gt $maxExtH) { $maxExtH = [int]$extHourCounts[$k] } }
        }
        foreach ($k in $intHourCounts.Keys) {
            if ($k.StartsWith("$s|")) { if ($intHourCounts[$k] -gt $maxIntH) { $maxIntH = [int]$intHourCounts[$k] } }
        }
        foreach ($k in $dayTotalCounts.Keys) {
            if ($k.StartsWith("$s|")) { if ($dayTotalCounts[$k] -gt $maxDay) { $maxDay = [int]$dayTotalCounts[$k] } }
        }

        [pscustomobject]@{
            SenderAddress               = $s
            Max_InternalRecipientsPerHr = $maxIntH
            Max_ExternalRecipientsPerHr = $maxExtH
            Max_TotalRecipientsPerDay   = $maxDay
            Rec_InternalPerHr           = [Math]::Max($MinInternalHourly, [int][Math]::Ceiling($Multiplier * $maxIntH))
            Rec_ExternalPerHr           = [Math]::Max($MinExternalHourly, [int][Math]::Ceiling($Multiplier * $maxExtH))
            Rec_TotalPerDay             = [Math]::Max($MinDaily,          [int][Math]::Ceiling($Multiplier * $maxDay))
        }
    }

    # Find the peak observed values across all senders (used to recommend global limits).
    $globalMaxIntH = if ($rowsOut.Count) { ($rowsOut | Measure-Object -Property Max_InternalRecipientsPerHr -Maximum).Maximum } else { 0 }
    $globalMaxExtH = if ($rowsOut.Count) { ($rowsOut | Measure-Object -Property Max_ExternalRecipientsPerHr -Maximum).Maximum } else { 0 }
    $globalMaxDay  = if ($rowsOut.Count) { ($rowsOut | Measure-Object -Property Max_TotalRecipientsPerDay -Maximum).Maximum } else { 0 }

    # Summary contains the recommended limits (multiplied + minimums) plus the observed peaks.
    $summary = [pscustomobject]@{
        HighestObserved_InternalPerHour = [int]$globalMaxIntH
        HighestObserved_ExternalPerHour = [int]$globalMaxExtH
        HighestObserved_TotalPerDay     = [int]$globalMaxDay
        Recommended_InternalPerHour     = [Math]::Max($MinInternalHourly, [int][Math]::Ceiling($Multiplier * $globalMaxIntH))
        Recommended_ExternalPerHour     = [Math]::Max($MinExternalHourly, [int][Math]::Ceiling($Multiplier * $globalMaxExtH))
        Recommended_TotalPerDay         = [Math]::Max($MinDaily,          [int][Math]::Ceiling($Multiplier * $globalMaxDay))
        Multiplier                      = $Multiplier
        DaysBack                        = $DaysBack
        ChunkDays                       = $ChunkDays
    }

    [pscustomobject]@{
        PerSender = ($rowsOut | Sort-Object Rec_TotalPerDay -Descending)
        Summary   = $summary
    }
}

# When this script is run directly, it will execute the function with a default set of parameters
# and write both the per-sender table and a summary to the console.
#
# You can uncomment / adjust the example below to target a specific set of sender domains:
# $validDomains = @("example.com","example.onmicrosoft.com","secondary.com")
# $result = Compute-MaxRecipientLimitsV2 -ValidSenderDomains $validDomains -DaysBack 30 -ChunkDays 5 -Multiplier 1.25

$result = Compute-MaxRecipientLimitsV2 -DaysBack 30 -ChunkDays 5 -Multiplier 1.25
$result.PerSender | Format-Table -AutoSize
"`nSUMMARY:"
$result.Summary | Format-List
