<#
.Synopsis
    Creates and emails user app audit reports to managers. Imports an Excel user listing and generates an Excel report for managers of 
    users listed in the Excel user listing. The report is then attached to an email asking the manager to confirm necessary access. 
    VIP Managers can be exempt by title or email, i.e. "Executive Vice President" or "CEO@company.com"
    This function requires an AD structure where Manager is set properly and a user report from the application with an email column. 
.DESCRIPTION
    Creates and emails user app audit reports to managers. Imports an Excel user listing and generates an Excel report for managers of 
    users listed in the Excel user listing. The report is then attached to an email asking the manager to confirm necessary access. 
    VIP Managers can be exempt by title or email, i.e. "Executive Vice President" or "CEO@company.com"
    This function requires an AD structure where Manager is set properly and a user report from the application with an email column.
.PARAMETER SubjectTemplate
    Name to give the subject of the emails and used within the body of the email. Typically "AppName User Review"
.PARAMETER ReportPath
    Path of the report to pick up i.e. ".\AppName_User_Report.xlsx"
.PARAMETER HeaderRow
    Row number the header of the xlsx is present on. Default is 1
.PARAMETER EmailColumnName
    Name of the column containing the user's email address. Default is "Email"
.PARAMETER Testing
    When true, sends all emails to the EmailTo address rather than the managers
.PARAMETER NoUserEmails
    When true, does not send emails to managers, only prints "would have sent" text
.PARAMETER EmailTo
    Email address to send reports that have no manager or are exceptions, as well as the main report listing
.PARAMETER EmailFrom
    Email address to send reports from.
.PARAMETER VIPTitle
    Title of managers who should not be sent emails. These emails will instead be sent to the EmailTo address
.PARAMETER VIPs
    Comma separated list of specific emails to exclude
.PARAMETER ReplyByDays
    Days out to request a reply from the manager by
.EXAMPLE
    New-AppAuditEmails -SubjectTemplate "Quickbooks User Review" -ReportPath ".\Quickbooks_User_Report.xlsx"
.EXAMPLE
    New-AppAuditEmails -SubjectTemplate "Quickbooks User Review" -ReportPath ".\Quickbooks_User_Report.xlsx" -Testing -NoUserEmails
.EXAMPLE
    New-AppAuditEmails -SubjectTemplate "Quickbooks User Review" -ReportPath ".\Quickbooks_User_Report.xlsx" -EmailColumnName "Email Address" -ReplyByDays 14
.EXAMPLE
    New-AppAuditEmails -SubjectTemplate "Quickbooks User Review" -ReportPath ".\Quickbooks_User_Report.xlsx" -VIPs "Accountant@company.com", "CEO@company.com"
#>
function New-AppAuditEmails
{
    [CmdletBinding()]
    [Alias()]
    Param
    (
        [Parameter(Mandatory=$true,Position=0)]
        [String]$SubjectTemplate,

        [Parameter(Mandatory=$true,Position=1)]
        [String]$ReportPath,

        [Parameter()]
        [int]$HeaderRow = 1,

        [Parameter()]
        [String]$EmailColumnName = "Email",

        [Parameter()]
        [switch]$Testing = $false,

        [Parameter()]
        [switch]$NoUserEmails = $false,

        #Parameters with defaults you will want to change
        [Parameter()]
        [String]$EmailTo = "ITSecurity@company.com",

        [Parameter()]
        [string]$EmailFrom = "ITSecurity@company.com",

        [Parameter()]
        [string]$VIPTitle = "Executive Vice President",

        [Parameter()]
        [string[]]$VIPs = @("VIP@company.com", "CEO@company.com", "CIO@company.com"),    

        [Parameter()]
        [int]$ReplyByDays = 7
    )

    Begin
    {
    }
    Process
    { 
        #May want to modify text
        [string]$BodyTemplate = 
    
            "Good morning, <br>
            <br>
            Information Security is conducting the semi-annual $($SubjectTemplate). 
            As part of our periodic review process, we reach out to application owners and managers in order to validate that user 
            permissions are appropriate for their current role or job responsibilities. Included with this email is the list of associates
            that have access to the application, along with the roles/permissions that have been provisioned to each employee. 
            The list of users that you are responsible for reviewing are included in the attached excel file. Please find the sheet titled
            with your name, There you can see the employees under you for reviewing. Respond to this email either affirming that the
            access for the employee is appropriate, or indicating any changes that need to be made to the existing access. Provide your 
            response by $($ReplyByDate). Changes based on the review performed will be requested and completed prior to the closure of 
            this process.
            <br>
            <br>
            Thanks,
            <br>
            <br>
            IT Security"

        $ReplyByDate = ((Get-Date).AddDays($ReplyByDays) | Get-Date -UFormat %D)
        
        $Report = @()
        $Report += Import-Excel -Path $ReportPath -HeaderRow $HeaderRow -DataOnly
         
        foreach ($Row in $Report) {
        [string]$Email = $Row.$EmailColumnName
        if($Email){
            $User = Get-ADUser -Filter {UserPrincipalName -like $Email} -Properties Manager
        }
        if ($User.Manager) {
                $ManagerEmail = (Get-ADUser -Identity $User.Manager).UserPrincipalName
                $ManagerName = (Get-ADUser -Identity $User.Manager).GivenName + " " + (Get-ADUser -Identity $User.Manager).Surname
                $ManagerTitle = (Get-ADUser -Identity $User.Manager -Properties extensionAttribute1).extensionAttribute1
        }
        else { 
                $ManagerEmail = "N/A"
                $ManagerName = "No Manager"
                $ManagerTitle = "N/A"
        }
        $Row | Add-Member -NotePropertyName "ManagerEmail" -NotePropertyValue $ManagerEmail
        $Row | Add-Member -NotePropertyName "ManagerName" -NotePropertyValue $ManagerName
        $Row | Add-Member -NotePropertyName "ManagerTitle" -NotePropertyValue $ManagerTitle
        }
    
        $Groups = $Report | Group-Object -Property ManagerName
        Foreach ($Group in $Groups) { 
            $GroupName = $Group.Name
            Export-Excel -InputObject $Group.Group -Path ".\$($GroupName).xlsx"
            
            
    
            $EmailSplat = @{
                SendTo = [string]::Empty
                Subject = ($SubjectTemplate)
                Body = ($BodyTemplate)
                Attachment = ".\$($GroupName).xlsx"
                SendFrom = $EmailFrom
            }
            if($Testing -eq $true){
                $EmailSplat['SendTo'] = $EmailTo
            }
            elseif($GroupName -eq "No Manager"){
                $EmailSplat['SendTo'] = $EmailTo
            }
            elseif($Group.Group[0].ManagerEmail -in $VIPs){
                $EmailSplat['SendTo'] = $EmailTo
            }
            elseif($Group.Group[0].ManagerTitle -eq $VIPTitle){
                $EmailSplat['SendTo'] = $EmailTo
            }
            else{
                $EmailSplat['SendTo'] = $Group.Group[0].ManagerEmail
            }
    
            if($NoUserEmails -eq $false){
                Write-Output "Sending $($EmailSplat.Attachment) to $($EmailSplat.SendTo)" `n
                Send-Email @EmailSplat
            }
            else{
                Write-Output "$($EmailSplat.Attachment) would have sent to $($EmailSplat.SendTo)" `n
            }
    
        } 
    
        #Send Full Report
        Export-Excel -InputObject $Report -Path ".\$($SubjectTemplate).xlsx" -AutoFilter
        $EmailSplat['SendTo'] = $EmailTo
        $EmailSplat['Attachment'] = ".\$($SubjectTemplate).xlsx"
        $EmailSplat['Body'] = "$($SubjectTemplate) attached."
        Send-Email @EmailSplat
    
        #Cleanup Workspace
        Remove-Item -Path ".\*.xlsx"
    
    }
    End
    {
    }
}