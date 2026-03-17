# The purpose of this script is to export a list of mailboxes from a Microsoft 365 tenant into a CSV file formatted for import into Autotask PSA as Configuration Items.
# This version authenticates interactively (no stored credential file) and will append to an existing output file without repeating the header row.

param(
    [string]$OutputPath,
    [string]$CompanyName,
    [string]$TenantId,
    [string]$UserPrincipalName
)

# Check if ExchangeOnlineManagement module is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Error: The 'ExchangeOnlineManagement' module is not installed. Please install it using: Install-Module -Name ExchangeOnlineManagement" -ForegroundColor Red
    exit
}

if (-not $CompanyName) {
    Write-Host "A CompanyName is required (used in the output for the 'Company' column)." -ForegroundColor Yellow
    Write-Host "Example: .\Export-MailboxesToAutotask.ps1 -CompanyName 'Contoso'" -ForegroundColor Yellow
    Exit 1
}

Write-Host "Connecting to Exchange Online (interactive sign-in) ..." -ForegroundColor Green
$connectParams = @{ ShowBanner = $false }
if ($TenantId) { $connectParams.TenantId = $TenantId }
if ($UserPrincipalName) { $connectParams.UserPrincipalName = $UserPrincipalName }

try {
    Connect-ExchangeOnline @connectParams

    $mailboxes = Get-Mailbox | Where-Object { $_.PrimarySmtpAddress -notmatch '^DiscoverySearchMailbox' }
    Start-Sleep -Seconds 2

    $outputObject = $mailboxes | ForEach-Object {
        Write-Host "Getting Mailbox Statistics for $($_.DisplayName) ..." -ForegroundColor Cyan
        $myMailboxStats = $_.Guid.Guid | Get-MailboxStatistics
        $myTotalItemSizeGB = $myMailboxStats.TotalItemSize -replace "(.*\()|,| [a-z]*\)", ""
        $myLastUserActionTime = $myMailboxStats.LastUserActionTime

        [PSCustomObject]@{
            '[required] Product Name' = 'Office 365 Mailbox'
            '[required] Company' = $CompanyName
            'Configuration Item Category' = 'Mailbox'
            'Configuration Item Type' = 'Exchange Mailbox'
            '[required] Install Date' = $_.WhenCreated
            'Serial Number' = $_.Guid
            'Reference Name' = $_.DisplayName
            'UDF:29683003 Primary Email Address' = $_.PrimarySmtpAddress
            'UDF:29683004 Import Date' = (Get-Date -Format d)
            'UDF:29683044 Mailbox Size [protected]' = $myTotalItemSizeGB
            'UDF:29683046 Mailbox Last User Action Time [protected]' = $myLastUserActionTime
            'UDF:29683047 Mail Forwarded To [protected]' = $_.ForwardingAddress
            'UDF:29683098 Mailbox Type [protected]' = $_.RecipientTypeDetails
        }
    }
}
catch {
    Write-Host "Failed to connect or retrieve mailboxes: $_" -ForegroundColor Red
    Exit 1
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
}

if (-not $OutputPath) {
    $OutputPath = Join-Path -Path (Get-Location) -ChildPath ("import_to_autotask_{0:yyyyMMdd}.csv" -f (Get-Date))
}

# Append to existing file without repeating the header row.
if (Test-Path $OutputPath) {
    $csvLines = $outputObject | ConvertTo-Csv -NoTypeInformation
    if ($csvLines.Count -gt 1) {
        $csvLines[1..($csvLines.Count - 1)] | Out-File -FilePath $OutputPath -Append -Encoding UTF8
    }
} else {
    $outputObject | Export-Csv -Path $OutputPath -NoTypeInformation
}

Write-Host "File written to $OutputPath" -ForegroundColor Green

# To get mailbox permissions, use $mailbox | Get-MailboxPermission | where-object {$_.IsInherited -eq $False -and $_.User -ne 'NT AUTHORITY\SELF' -and $_.User -notlike 'S-1*'}
# To get forwarding address, use $mailbox.ForwardingAddress
# Other goodies exist in the Mailbox Statistics like ItemCount and LastLogonTime.
