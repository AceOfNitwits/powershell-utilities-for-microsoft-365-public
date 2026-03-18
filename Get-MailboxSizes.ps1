#Requires -Modules ExchangeOnlineManagement

Connect-ExchangeOnline -ShowBanner:$false

function Get-BytesFromTotalItemSize {
    param(
        [Parameter(Mandatory)]
        $TotalItemSize
    )

    # Typically looks like: "1.234 GB (1,325,123,456 bytes)"
    $s = $TotalItemSize.ToString()

    if ($s -match '\(([0-9,]+)\sbytes\)') {
        return [int64]($matches[1] -replace ',', '')
    }

    return $null
}

$results = foreach ($mbx in Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited) {

    $stats = Get-MailboxStatistics -Identity $mbx.UserPrincipalName

    $primaryBytes = Get-BytesFromTotalItemSize -TotalItemSize $stats.TotalItemSize
    $primarySizeGB = if ($null -ne $primaryBytes) { [math]::Round($primaryBytes / 1GB, 2) } else { $null }

    $archiveSizeGB = $null
    if ($mbx.ArchiveStatus -eq "Active") {
        $archiveStats = Get-MailboxStatistics -Identity $mbx.UserPrincipalName -Archive
        $archiveBytes = Get-BytesFromTotalItemSize -TotalItemSize $archiveStats.TotalItemSize
        $archiveSizeGB = if ($null -ne $archiveBytes) { [math]::Round($archiveBytes / 1GB, 2) } else { $null }
    }

    [pscustomobject]@{
        DisplayName       = $mbx.DisplayName
        UserPrincipalName = $mbx.UserPrincipalName
        PrimarySizeGB     = $primarySizeGB
        PrimaryItemCount  = $stats.ItemCount
        ArchiveSizeGB     = $archiveSizeGB
    }
}

# Screen output (largest first)
$results | Sort-Object PrimarySizeGB -Descending | Format-Table -AutoSize

Disconnect-ExchangeOnline -Confirm:$false