#Requires -Version 3.0

# Always load WPF assembly to be able to use "[System.Windows.MessageBox]"
Add-Type -AssemblyName presentationframework, presentationcore

# a message, a title and a button
# More info : https://msdn.microsoft.com/en-us/library/ms598690.aspx
$msg = "WARNING: This script is collecting ALL mailboxes server by server, and ALL MAilbox statistics. Click [CANCEL] to abort, [OK] to continue"
$Title = "Continue ?"
$Button = "OkCancel"
$Result = [System.Windows.MessageBox]::Show($msg,$Title, $Button)
If ($Result -eq "Cancel"){Write-Host "You chose to cancel.";exit}Else{Write-Host "You chose to continue, gathering Exchange and mailbox statistics information..."}


Add-PSSnapin microsoft.exchange.management.powershell.e2010

# Setting variable section
$Partner = "CUSTOMER_NAME"
$CurrentDate = Get-Date
$CurrentDate = $CurrentDate.ToString('MM-dd-yyyy_hh-mm-ss')
$ExportFilePath = "c:\temp\Mailbox_Report_$CurrentDate.csv"
#End setting variable section

$MailboxServers = Get-MailboxServer
$Collection = @()

Foreach ($Server in $MailboxServers){
        $Mailboxes = Get-Mailbox -Server $Server -ResultSize Unlimited
        Foreach ($Mailbox in $MAilboxes) {
            # Putting properties we're interested about in separate variables (that's not necessary, but it can help making things a bit clearer)
            $DisplayName = $Mailbox.DisplayName
            $email = $Mailbox.primarysmtpaddress
            $MailboxQuota = $Mailbox.prohibitSendReceiveQuota.Value
            
            If ($MAilboxQuota -eq 0 -or $MAilboxQuota -eq $null){$MailboxQuota = 0}Else{$MailboxQuota = $Mailbox.prohibitSendReceiveQuota.value.ToBytes()}
            
            #$database = $Mailbox.database
            $Database = Get-MailboxDatabase $($Mailbox.Database)
            $DatabaseQuota = $Database.ProhibitSendReceiveQuota.Value.ToBytes()
            $UseDatabaseQuotaDefault = $Mailbox.UseDatabaseQuotaDefaults
            $type = $Mailbox.RecipientTypeDetails
            If ($UseDatabaseQuotaDefault -eq $true){$Limit = $DatabaseQuota} Else {$Limit = $MailboxQuota}
            #At this point, the $Limit will be the Database Quota if mailbox is using Database Quota defaults, and
            # $limit will be the Mailbox quota if mailbox is NOT using database quota defaults

            #We can call Get-MailboxStatistics via the Foreach $Mailbox variable, or directly on the $email variable where we stored the primary SMTP address a few lines above
            $CurrentMailboxStats = Get-MailboxStatistics -Identity $Mailbox.Identity| Select TotalItemSize
            $TotalMailboxSize = $CurrentMailboxStats.Totalitemsize.VAlue.ToBytes()

            If ($TotalMailboxSize -ge $($Limit * 0.85)){$Status = "WARNING"}ELSE{$STATUS = "OK"}

            $CurrentCustomObject = [PSCustomObject]@{
                                            "Partner" = $Partner
                                            "Display Name" = $DisplayName
                                            "Email Address" = $email
                                            Type = $type
                                            "Current Mailbox Size in GB" = [math]::round($TotalMailboxSize/1GB,2)
                                            "Mailbox Size Limit in GB" = [math]::round($Limit/1GB,2)
                                            Status = $Status
                                            "Use Database Quotas" = $UseDatabaseQuotaDefault
                                            "Database Quota" = [math]::round($DatabaseQuota/1GB,2)
                                            "MailboxQuota" = [math]::round($MAilboxQuota/1GB,2)
                                       }
             $Collection += $CurrentCustomObject

           }
}


$Collection | Export-CSV $ExportFilePath -NoTypeInformation -Encoding UTF8

Notepad $ExportFilePath
