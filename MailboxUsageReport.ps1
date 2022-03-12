<#
.SYNOPSIS
  Generate a report about mailboxes population , usage , qouta , and last logon for Exchange on-prem.

.DESCRIPTION
  This script will prompt for CSV file to store the report

.INPUTS
  None

.OUTPUTS
  CSV file

.EXAMPLE
  .\MailboxUsageReport.ps1
#>

Param(
    [Parameter(Mandatory = $false)]
    [SWITCH] $FixMbx
)

$ReportPathCSV =.\MailboxUsageReport.CSV

# You might need to use this switch to fix inaccessible mailboxes or mailboxes that were created in previous versions of Exchange.

if($FixMbx) {
  Write-Verbose -Message "Fixing inaccessible mailboxes or mailboxes that were created in previous versions of Exchange" -Verbose
  Get-Mailbox -resultsize unlimited |Set-Mailbox -ApplyMandatoryProperties }

$Mailboxes = Get-Mailbox -ResultSize Unlimited
$Total=$Mailboxes.count
Write-host "Total No. of Mailboxes Is="$Total
$c=0
$Report = @()
foreach($Mailbox in $Mailboxes)
    {
        $MailboxStats = (Get-MailboxStatistics $Mailbox -WarningAction SilentlyContinue )

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
        $ReportItem | Add-Member -Type NoteProperty -Name DisplayName -Value $Mailbox.DisplayName
        $ReportItem | Add-Member -Type NoteProperty -Name UserName -Value $Mailbox.SamAccountName
        $ReportItem | Add-Member -Type NoteProperty -Name PrimarySMTP -Value $Mailbox.WindowsEmailAddress
        $ReportItem | Add-Member -Type NoteProperty -Name OrganizationalUnit -Value $Mailbox.OrganizationalUnit
        $ReportItem | Add-Member -Type NoteProperty -Name Database -Value $Mailbox.Database
        $ReportItem | Add-Member -Type NoteProperty -Name ExchangeUserAccountControl -Value $Mailbox.ExchangeUserAccountControl
	    $ReportItem | Add-Member -Type NoteProperty -Name RecipientTypeDetailst -Value $Mailbox.RecipientTypeDetails
        $ReportItem | Add-Member -Type NoteProperty -Name Accountdisabled -Value $Mailbox.Accountdisabled
        $ReportItem | Add-Member -Type NoteProperty -Name WhenCreated -Value $Mailbox.WhenCreated
        $ReportItem | Add-Member -Type NoteProperty -Name ArchiveState -Value $Mailbox.ArchiveState
        $ReportItem | Add-Member -Type NoteProperty -Name ArchiveQuota -Value $Mailbox.ArchiveQuota
        $ReportItem | Add-Member -Type NoteProperty -Name EmailAliases -Value ($Mailbox.EmailAddresses.SmtpAddress -join "; ")
        if($MailboxStats)
            {
                $ReportItem | Add-Member -Type NoteProperty -Name LastLogonTime -Value $MailboxStats.LastLogonTime
                $ReportItem | Add-Member -Type NoteProperty -Name TotalItemSizeInMB -Value $MailboxStats.TotalItemSize.Value.ToMB()
                $ReportItem | Add-Member -Type NoteProperty -Name ItemCount -Value $MailboxStats.ItemCount
                $ReportItem | Add-Member -Type NoteProperty -Name DeletedItemCount -Value $MailboxStats.DeletedItemCount
                $ReportItem | Add-Member -Type NoteProperty -Name TotalDeletedItemSizeInMB -Value $MailboxStats.TotalDeletedItemSize.Value.ToMB()

            }
        $ReportItem | Add-Member -Type NoteProperty -Name StorageLimitStatus -Value $Mailbox.StorageLimitStatus
        $ReportItem | Add-Member -Type NoteProperty -Name UseDatabaseQuotaDefaults -Value $Mailbox.UseDatabaseQuotaDefaults
        $ReportItem | Add-Member -Type NoteProperty -Name ProhibitSendReceiveQuotaInMB -Value $ProhibitSendReceiveQuota
        $Report += $ReportItem
    }
write-host "The Output is .\MailboxUsageReport.CSV File"
$Report | Sort TotalItemSize -Descending | Export-Csv -NoTypeInformation $ReportPathCSV