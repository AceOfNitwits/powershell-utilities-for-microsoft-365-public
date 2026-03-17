# As stored credentials are profile specific, the credentials file will be stored in the user profile folder.
$tennantListPath = "$env:USERPROFILE\Microsoft365CompanyCreds.csv"
try {
    $tennants = Import-Csv $tennantListPath # Import the credential file so we can iterate through it.
}
catch {
    Write-Host "List of tennants not found. Please create a .csv file in the format CompanyName, UserName, Password (where the password is plain-text encrypted) in the following path: $tennantListPath" -ForegroundColor Red
    Exit    
}
[PSCustomObject]$outputObject =@()
$tennants | ForEach-Object { # Iterate through the credential file.
    $myMailboxes = @() #Initialize the variable so we don't get carryover from previous tennants if we cannot connect.
    [string]$companyName = $_.CompanyName
    [string]$userName = $_.UserName
    Write-Host "Connecting to $companyName as $userName`." -ForegroundColor Black -backgroundcolor Green
    [securestring]$securePwd = $_.Password | ConvertTo-SecureString
    [PSCredential]$UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName, $securePwd
    # $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    #Import-PSSession $Session
    try {
        Connect-ExchangeOnline -Credential $UserCredential -ShowBanner:$false
        # The $myOutput variable will hold the collection of custom objects created by the ForEach loop through the mailboxes.
        $myMailboxes = Get-Mailbox | Where-Object {$_.primarysmtpaddress -notmatch '^DiscoverySearchMailbox'}
        Start-Sleep -seconds 2 # If we try to run other queries too quickly after the Get-Mailbox, we get errors.
        [pscustomobject]$myOutput = $myMailboxes | ForEach-Object {
            #Start-Sleep -Seconds 2 # If we try to run other queries too quickly after the Get-Mailbox, we get errors. This is a terrible solution though.
            Write-Host "Getting Mailbox Statistics for $($_.DisplayName)`."
            $myMailboxStats = $_.guid.guid | Get-MailboxStatistics
            $myTotalItemSizeGB = $myMailboxStats.TotalItemSize -replace “(.*\()|,| [a-z]*\)”, “”
            $myLastUserActionTime = $myMailboxStats.LastUserActionTime
    <#        Write-Host "Getting Folder Statistics for $($_.DisplayName)`."
            $myMailboxFolderStats = $_ | Get-MailboxFolderStatistics
            $myFolderCount = $myMailboxFolderStats | Measure-Object | Select-Object -ExpandProperty Count
    #>
            [PSCustomObject]@{
                '[required] Product Name' = 'Office 365 Mailbox'
                '[required] Company' = $companyName
                'Configuration Item Category' = 'Exchange Mailbox'
                'Configuration Item Type' = 'Exchange Mailbox'
                '[required] Install Date' = $_.WhenCreated
                'Serial Number' = $_.Guid
                'Reference Name' = $_.DisplayName
                'UDF:29683003 Primary Email Address' = $_.PrimarySmtpAddress
                'UDF:29683004 Import Date' = (get-date -Format d)
                'UDF:29683044 Mailbox Size [protected]' = $myTotalItemSizeGB
                #'UDF:29683045 Mailbox Folder Count [protected]' = $myFolderCount
                'UDF:29683046 Mailbox Last User Action Time [protected]' = $myLastUserActionTime
                'UDF:29683047 Mail Forwarded To [protected]' = $_.ForwardingAddress
                'UDF:29683098 Mailbox Type [protected]' = $_.RecipientTypeDetails
            } 
        }
    
        # Remove-PSSession $Session # Close out our session.
        Disconnect-ExchangeOnline -Confirm:$false
        $outputObject += $myOutput # Add the current collection of objects to the master collection.
    } catch {
        # There was a problem.
    }
}
$outputObject | format-table -AutoSize
$outputPath = "$env:USERPROFILE\OneDrive - Clocktower Technology Services, Inc\Documents\Scripts\M365-Management\import_to_autotask_$(get-date -Format 'yyyyMMddTHHmmss').csv"
$outputObject | Export-Csv -Path ($outputPath) -NoTypeInformation # Export the collection of objects to a CSV file suitable for import into Autotask.
Write-Host "File written to $outputPath"

# To get mailbox permissions, use $mailbox | Get-MailboxPermission | where-object {$_.IsInherited -eq $False -and $_.User -ne 'NT AUTHORITY\SELF' -and $_.User -notlike 'S-1*'}
# To get forwarding address, use $mailbox.ForwardingAddress
# Other goodies exist in the Mailbox Statistics like ItemCount and LastLogonTime.