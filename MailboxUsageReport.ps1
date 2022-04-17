<#
.SYNOPSIS
  Generate a report about mailboxes population , usage , qouta , and last logon for Exchange on-prem.

.DESCRIPTION
  This script will prompt for CSV file to store the report

.INPUTS
  None

.OUTPUTS
  CSV file

.EXAMPLE 1
  .\MailboxUsageReport.ps1
.EXAMPLE 2 :you might need to use (-FixMbx) switch to fix inaccessible mailboxes or mailboxes that were created in previous versions of Exchange.
  .\MailboxUsageReport.ps1 -FixMbx
#>

Param(
    [Parameter(Mandatory = $false)]
    [SWITCH] $FixMbx
)

$ReportFileName =".\MailboxUsageReport.CSV"
$DataFileName =".\MaskedData.csv"

# You might need to use this switch to fix inaccessible mailboxes or mailboxes that were created in previous versions of Exchange.

if($FixMbx) {
  Write-Verbose -Message "Fixing inaccessible mailboxes or mailboxes that were created in previous versions of Exchange" -Verbose
  Get-Mailbox -resultsize unlimited -IgnoreDefaultScope|Set-Mailbox -ApplyMandatoryProperties }

$Mailboxes = Get-Mailbox -ResultSize Unlimited -IgnoreDefaultScope
$Total=$Mailboxes.count
Write-host "Total No. of Mailboxes Is="$Total
$c=0
$Report = @()
$MaskedData = @()
foreach($Mailbox in $Mailboxes)
    {
        $MailboxStats = (Get-MailboxStatistics $Mailbox -WarningAction SilentlyContinue )
        $User = (Get-user $Mailbox.userprincipalname -WarningAction SilentlyContinue )

        if ($mailbox.UseDatabaseQuotaDefaults -eq $true)
            {
                $ProhibitSendReceiveQuota = (Get-MailboxDatabase $mailbox.Database).ProhibitSendReceiveQuota.Value.ToMB()
            }
        if ($mailbox.UseDatabaseQuotaDefaults -eq $false)
            {
                $ProhibitSendReceiveQuota = $mailbox.ProhibitSendReceiveQuota.Value.ToMB()
            }
        $c++
        Write-host "Working on Mailbox "$c "Out of "$Total
        $ReportItem = New-Object System.Object
        $MaskedDataItem = New-Object System.Object
        $ReportItem | Add-Member -Type NoteProperty -Name DisplayName -Value $Mailbox.DisplayName
        $ReportItem | Add-Member -Type NoteProperty -Name UserName -Value $Mailbox.SamAccountName
        $ReportItem | Add-Member -Type NoteProperty -Name PrimarySMTP -Value $Mailbox.WindowsEmailAddress
        $ReportItem | Add-Member -Type NoteProperty -Name OrganizationalUnit -Value $Mailbox.OrganizationalUnit
        $MaskedDataItem |Add-Member -Type NoteProperty -Name OrganizationalUnit -Value $Mailbox.OrganizationalUnit
        $ReportItem | Add-Member -Type NoteProperty -Name Database -Value $Mailbox.Database
        $MaskedDataItem | Add-Member -Type NoteProperty -Name Database -Value $Mailbox.Database
        $ReportItem | Add-Member -Type NoteProperty -Name ExchangeUserAccountControl -Value $Mailbox.ExchangeUserAccountControl
        $MaskedDataItem |Add-Member -Type NoteProperty -Name ExchangeUserAccountControl -Value $Mailbox.ExchangeUserAccountControl
	    $ReportItem | Add-Member -Type NoteProperty -Name RecipientTypeDetailst -Value $Mailbox.RecipientTypeDetails
        $MaskedDataItem | Add-Member -Type NoteProperty -Name RecipientTypeDetailst -Value $Mailbox.RecipientTypeDetails
        $ReportItem | Add-Member -Type NoteProperty -Name Accountdisabled -Value $Mailbox.Accountdisabled
        $MaskedDataItem | Add-Member -Type NoteProperty -Name Accountdisabled -Value $Mailbox.Accountdisabled
        $ReportItem | Add-Member -Type NoteProperty -Name WhenCreated -Value $Mailbox.WhenCreated
        $MaskedDataItem | Add-Member -Type NoteProperty -Name WhenCreated -Value $Mailbox.WhenCreated
        $ReportItem | Add-Member -Type NoteProperty -Name ArchiveState -Value $Mailbox.ArchiveState
        $MaskedDataItem | Add-Member -Type NoteProperty -Name ArchiveState -Value $Mailbox.ArchiveState
        $ReportItem | Add-Member -Type NoteProperty -Name ArchiveQuota -Value $Mailbox.ArchiveQuota
        $MaskedDataItem | Add-Member -Type NoteProperty -Name ArchiveQuota -Value $Mailbox.ArchiveQuota
        $ReportItem | Add-Member -Type NoteProperty -Name LitigationHoldEnabled -Value $Mailbox.LitigationHoldEnabled
        $MaskedDataItem | Add-Member -Type NoteProperty -Name LitigationHoldEnabled -Value $Mailbox.LitigationHoldEnabled
        $ReportItem | Add-Member -Type NoteProperty -Name RetentionHoldEnabled -Value $Mailbox.RetentionHoldEnabled
        $MaskedDataItem | Add-Member -Type NoteProperty -Name RetentionHoldEnabled -Value $Mailbox.RetentionHoldEnabled
        $ReportItem | Add-Member -Type NoteProperty -Name SingleItemRecoveryEnabled -Value $Mailbox.SingleItemRecoveryEnabled
        $MaskedDataItem | Add-Member -Type NoteProperty -Name SingleItemRecoveryEnabled -Value $Mailbox.SingleItemRecoveryEnabled
        $ReportItem | Add-Member -Type NoteProperty -Name Department -Value $User.Department
        $MaskedDataItem | Add-Member -Type NoteProperty -Name Department -Value $User.Department
        $ReportItem | Add-Member -Type NoteProperty -Name City -Value $User.City
        $MaskedDataItem | Add-Member -Type NoteProperty -Name City -Value $User.City
        $ReportItem | Add-Member -Type NoteProperty -Name StateOrProvince -Value $User.StateOrProvince
        $MaskedDataItem | Add-Member -Type NoteProperty -Name StateOrProvince -Value $User.StateOrProvince
        $ReportItem | Add-Member -Type NoteProperty -Name CountryOrRegion -Value $User.CountryOrRegion
        $MaskedDataItem | Add-Member -Type NoteProperty -Name CountryOrRegion -Value $User.CountryOrRegion
        $ReportItem | Add-Member -Type NoteProperty -Name EmailAliases -Value ($Mailbox.EmailAddresses.SmtpAddress -join "; ")
        
        if($MailboxStats)
            {
                $ReportItem | Add-Member -Type NoteProperty -Name LastLogonTime -Value $MailboxStats.LastLogonTime
                $MaskedDataItem | Add-Member -Type NoteProperty -Name LastLogonTime -Value $MailboxStats.LastLogonTime
                $ReportItem | Add-Member -Type NoteProperty -Name TotalItemSizeInMB -Value $MailboxStats.TotalItemSize.Value.ToMB()
                $MaskedDataItem | Add-Member -Type NoteProperty -Name TotalItemSizeInMB -Value $MailboxStats.TotalItemSize.Value.ToMB()
                $ReportItem | Add-Member -Type NoteProperty -Name ItemCount -Value $MailboxStats.ItemCount
                $MaskedDataItem | Add-Member -Type NoteProperty -Name ItemCount -Value $MailboxStats.ItemCount
                $ReportItem | Add-Member -Type NoteProperty -Name DeletedItemCount -Value $MailboxStats.DeletedItemCount
                $MaskedDataItem | Add-Member -Type NoteProperty -Name DeletedItemCount -Value $MailboxStats.DeletedItemCount
                $ReportItem | Add-Member -Type NoteProperty -Name TotalDeletedItemSizeInMB -Value $MailboxStats.TotalDeletedItemSize.Value.ToMB()
                $MaskedDataItem | Add-Member -Type NoteProperty -Name TotalDeletedItemSizeInMB -Value $MailboxStats.TotalDeletedItemSize.Value.ToMB()

            }
        
        $ReportItem | Add-Member -Type NoteProperty -Name UseDatabaseQuotaDefaults -Value $Mailbox.UseDatabaseQuotaDefaults
        $MaskedDataItem | Add-Member -Type NoteProperty -Name UseDatabaseQuotaDefaults -Value $Mailbox.UseDatabaseQuotaDefaults
        $ReportItem | Add-Member -Type NoteProperty -Name ProhibitSendReceiveQuotaInMB -Value $ProhibitSendReceiveQuota
        $MaskedDataItem | Add-Member -Type NoteProperty -Name ProhibitSendReceiveQuotaInMB -Value $ProhibitSendReceiveQuota
        $Report += $ReportItem
        $MaskedData += $MaskedDataItem
    }
write-host "The Report File saved in .\MailboxUsageReport.CSV"
write-host "Masked Data File Saved in .\UsageMaskedData.CSV"
$Report | Export-Csv -NoTypeInformation $ReportFileName
$MaskedData | Export-Csv -NoTypeInformation $DataFileName