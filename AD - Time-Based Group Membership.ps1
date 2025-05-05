#########
# au2mator PS Services
# Type: New Service
#
# Title: AD - Time-Based Group Membership
#
# v 1.0 Initial Release
# 
#
# Init Release: 08.08.2021
# Last Update: 08.08.2021
# Code Template V 1.4
#
# URL: https://au2mator.com/time-based-group-membership-active-directory-self-service-with-au2mator/?utm_source=github&utm_medium=social&utm_campaign=AD_TempGroupMembership&utm_content=PS1
# Github: https://github.com/au2mator/AD-Time-Based-Group-Membership
#################

#region InputParamaters
##Question in au2mator

param (
    [parameter(Mandatory = $false)] 
    [String]$c_User,

    [parameter(Mandatory = $false)] 
    [String]$c_Group,

    [parameter(Mandatory = $false)] 
    [String]$c_Type,

    [parameter(Mandatory = $false)] 
    [String]$c_EndDate,


    [parameter(Mandatory = $false)] 
    [String]$c_TTL,


    [parameter(Mandatory = $true)] 
    [String]$InitiatedBy, 

    [parameter(Mandatory = $true)] 
    [String]$RequestId, 
 
    [parameter(Mandatory = $true)] 
    [String]$Service, 
 
    [parameter(Mandatory = $true)] 
    [String]$TargetUserId
)

#endregion  InputParamaters


#region Variables
Set-ExecutionPolicy -ExecutionPolicy Bypass
$DoImportPSSession = $false


## Environment
[string]$PSRemotingServer = "demo01"
[string]$LogPath = "C:\_SCOworkingDir\TFS\PS-Services\AD - Time-Based (Temporary) Group Membership"
[string]$LogfileName = "AD - Time-Based Group Membership"

[string]$CredentialStorePath = "C:\_SCOworkingDir\TFS\PS-Services\CredentialStore" #see for details: https://click.au2mator.com/PSCreds/?utm_source=github&utm_medium=social&utm_campaign=AD_TempGroupMembership&utm_content=PS1



$Modules = @("ActiveDirectory") #$Modules = @("ActiveDirectory", "SharePointPnPPowerShellOnline")


## au2mator Settings
[string]$PortalURL = "http://demo01.au2mator.local"
[string]$au2matorDBServer = "demo01"
[string]$au2matorDBName = "au2mator40Demo2"


#Control Teams
$SendTeamsCardToInitiatedByUser = $true #Send a Card in Teams after Service is completed
$SendTeamsCardToTargetUser = $true #Send Card in Teams to Target User after Service is completed
$ToChannel = $true #Send the Message Card to a Channel
$ToUser = $false #Send the Message Card to the User via Chat

#Teams Settings
$TeamName = "au2mator - ORG"
$ChannelName = "General"

## Control Mail
$SendMailToInitiatedByUser = $true #Send a Mail after Service is completed
$SendMailToTargetUser = $false #Send Mail to Target User after Service is completed

## SMTP Settings
$SMTPServer = "smtp.office365.com"
$SMPTAuthentication = $true #When True, User and Password needed
$EnableSSLforSMTP = $true
$SMTPSender = "SelfService@au2mator.com"
$SMTPPort = "587"

# Stored Credentials
# See: https://click.au2mator.com/PSCreds/?utm_source=github&utm_medium=social&utm_campaign=AD_TempGroupMembership&utm_content=PS1
$SMTPCredential_method = "Stored" #Stored, Manual
$SMTPcredential_File = "SMTPCreds.xml"
$SMTPUser = ""
$SMTPPassword = ""

if ($SMTPCredential_method -eq "Stored") {
    $SMTPcredential = Import-CliXml -Path (Get-ChildItem -Path $CredentialStorePath -Filter $SMTPcredential_File).FullName
}

if ($SMTPCredential_method -eq "Manual") {
    $f_secpasswd = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
    $SMTPcredential = New-Object System.Management.Automation.PSCredential ($SMTPUser, $f_secpasswd)
}

#endregion Variables


#region CustomVariables

$GraphAPICred_File = "TeamsCreds.xml"
$GraphAPICred = Import-CliXml -Path (Get-ChildItem -Path $CredentialStorePath -Filter $GraphAPICred_File).FullName
$clientId = $GraphAPICred.clientId
$clientSecret = $GraphAPICred.clientSecret
$tenantName = $GraphAPICred.tenantName
$User = $GraphAPICred.User
$PW = $GraphAPICred.PW



#endregion CustomVariables




#region Functions
function Write-au2matorLog {
    [CmdletBinding()]
    param
    (
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR')]
        [string]$Type,
        [string]$Text
    )

    # Set logging path
    if (!(Test-Path -Path $logPath)) {
        try {
            $null = New-Item -Path $logPath -ItemType Directory
            Write-Verbose ("Path: ""{0}"" was created." -f $logPath)
        }
        catch {
            Write-Verbose ("Path: ""{0}"" couldn't be created." -f $logPath)
        }
    }
    else {
        Write-Verbose ("Path: ""{0}"" already exists." -f $logPath)
    }
    [string]$logFile = '{0}\{1}_{2}.log' -f $logPath, $(Get-Date -Format 'yyyyMMdd'), $LogfileName
    $logEntry = '{0}: <{1}> <{2}> <{3}> {4}' -f $(Get-Date -Format dd.MM.yyyy-HH:mm:ss), $Type, $RequestId, $Service, $Text
    Add-Content -Path $logFile -Value $logEntry
}

function ConnectToDB {
    # define parameters
    param(
        [string]
        $servername,
        [string]
        $database
    )
    Write-au2matorLog -Type INFO -Text "Function ConnectToDB"
    # create connection and save it as global variable
    $global:Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = "server='$servername';database='$database';trusted_connection=false; integrated security='true'"
    $Connection.Open()
    Write-au2matorLog -Type INFO -Text 'Connection established'
}

function ExecuteSqlQuery {
    # define parameters
    param(

        [string]
        $sqlquery

    )
    Write-au2matorLog -Type INFO -Text "Function ExecuteSqlQuery"
    #Begin {
    If (!$Connection) {
        Write-au2matorLog -Type WARNING -Text"No connection to the database detected. Run command ConnectToDB first."
    }
    elseif ($Connection.State -eq 'Closed') {
        Write-au2matorLog -Type INFO -Text 'Connection to the database is closed. Re-opening connection...'
        try {
            # if connection was closed (by an error in the previous script) then try reopen it for this query
            $Connection.Open()
        }
        catch {
            Write-au2matorLog -Type INFO -Text "Error re-opening connection. Removing connection variable."
            Remove-Variable -Scope Global -Name Connection
            Write-au2matorLog -Type WARNING -Text "Unable to re-open connection to the database. Please reconnect using the ConnectToDB commandlet. Error is $($_.exception)."
        }
    }
    #}

    #Process {
    #$Command = New-Object System.Data.SQLClient.SQLCommand
    $command = $Connection.CreateCommand()
    $command.CommandText = $sqlquery

    Write-au2matorLog -Type INFO -Text "Running SQL query '$sqlquery'"
    try {
        $result = $command.ExecuteReader()
    }
    catch {
        $Connection.Close()
    }
    $Datatable = New-Object "System.Data.Datatable"
    $Datatable.Load($result)

    return $Datatable

    #}

    #End {
    Write-au2matorLog -Type INFO -Text "Finished running SQL query."
    #}
}

function Get-UserInput ($RequestID) {
    [hashtable]$return = @{ }

    Write-au2matorLog -Type INFO -Text "Function Get-UserInput"
    ConnectToDB -servername $au2matorDBServer -database $au2matorDBName

    $Result = ExecuteSqlQuery -sqlquery "SELECT        RPM.Text AS Question, RP.Value
    FROM            dbo.Requests AS R INNER JOIN
                             dbo.RunbookParameterMappings AS RPM ON R.ServiceId = RPM.ServiceId INNER JOIN
                             dbo.RequestParameters AS RP ON RPM.ParameterName = RP.[Key] AND R.ID = RP.RequestId
    where RP.RequestId = '$RequestID' and rpm.IsDeleted = '0' order by [Order]"

    $html = "<table><tr><td><b>Question</b></td><td><b>Answer</b></td></tr>"
    $html = "<table>"
    foreach ($row in $Result) {
        #$row
        $html += "<tr><td><b>" + $row.Question + ":</b></td><td>" + $row.Value + "</td></tr>"
    }
    $html += "</table>"

    $f_RequestInfo = ExecuteSqlQuery -sqlquery "select InitiatedBy, TargetUserId,[ApprovedBy], [ApprovedTime], Comment from Requests where Id =  '$RequestID'"

    $Connection.Close()
    Remove-Variable -Scope Global -Name Connection

    $f_SamInitiatedBy = $f_RequestInfo.InitiatedBy.Split("\")[1]
    $f_UserInitiatedBy = Get-ADUser -Identity $f_SamInitiatedBy -Properties Mail


    $f_SamTarget = $f_RequestInfo.TargetUserId.Split("\")[1]
    $f_UserTarget = Get-ADUser -Identity $f_SamTarget -Properties Mail

    $return.InitiatedBy = $f_RequestInfo.InitiatedBy.trim()
    $return.MailInitiatedBy = $f_UserInitiatedBy.mail.trim()
    $return.MailTarget = $f_UserTarget.mail.trim()
    $return.TargetUserId = $f_RequestInfo.TargetUserId.trim()
    $return.ApprovedBy = $f_RequestInfo.ApprovedBy.trim()
    $return.ApprovedTime = $f_RequestInfo.ApprovedTime
    $return.Comment = $f_RequestInfo.Comment
    $return.HTML = $HTML

    return $return
}

Function Get-MailContent ($RequestID, $RequestTitle, $EndDate, $TargetUserId, $InitiatedBy, $Status, $PortalURL, $RequestedBy, $AdditionalHTML, $InputHTML) {

    Write-au2matorLog -Type INFO -Text "Function Get-MailContent"
    $f_RequestID = $RequestID
    $f_InitiatedBy = $InitiatedBy

    $f_RequestTitle = $RequestTitle

    try {
        $f_EndDate = (get-Date -Date $EndDate -Format (Get-Culture).DateTimeFormat.ShortDatePattern) + " (" + (get-Date -Date $EndDate -Format (Get-Culture).DateTimeFormat.ShortTimePattern) + ")"
    }
    catch {
        $f_EndDate = $EndDate
    }

    $f_RequestStatus = $Status
    $f_RequestLink = "$PortalURL/requeststatus?id=$RequestID"
    $f_HTMLINFO = $AdditionalHTML
    $f_InputHTML = $InputHTML

    $f_SamInitiatedBy = $f_InitiatedBy.Split("\")[1]
    $f_UserInitiatedBy = Get-ADUser -Identity $f_SamInitiatedBy -Properties DisplayName
    $f_DisplaynameInitiatedBy = $f_UserInitiatedBy.DisplayName


    $HTML = @'
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 1.5pt; background: #F7F8F3; mso-yfti-tbllook: 1184;" border="0" width="100%" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;">
    <td style="padding: .75pt .75pt .75pt .75pt;" valign="top">&nbsp;</td>
    <td style="width: 450.0pt; padding: .75pt .75pt .75pt .75pt; box-sizing: border-box;" valign="top" width="600">
    <div style="box-sizing: border-box;">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: white; border: solid #E9E9E9 1.0pt; mso-border-alt: solid #E9E9E9 .75pt; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="1" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes;">
    <td style="border: none; background: #6ddc36; padding: 15.0pt 0cm 15.0pt 15.0pt;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><img src="https://au2mator.com/wp-content/uploads/2018/02/HPLogoau2mator-1.png" alt="" width="198" height="43" /></p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 1; box-sizing: border-box;">
    <td style="border: none; padding: 15.0pt 15.0pt 15.0pt 15.0pt; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="width: 55.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top" width="55%">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes;">
    <td style="width: 18.75pt; border-top: solid #E3E3E3 1.0pt; border-left: solid #E3E3E3 1.0pt; border-bottom: none; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" width="25">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: none; border-bottom: none; border-right: solid #E3E3E3 1.0pt; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm; font-color: #0000;"><strong>End Date</strong>: ##EndDate</td>
    </tr>
    <tr style="mso-yfti-irow: 1;">
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: solid #E3E3E3 1.0pt; border-bottom: none; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 0cm 0cm;">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border-top: solid #E3E3E3 1.0pt; border-left: none; border-bottom: none; border-right: solid #E3E3E3 1.0pt; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;"><strong>Status</strong>: ##Status</td>
    </tr>
    <tr style="mso-yfti-irow: 2; mso-yfti-lastrow: yes;">
    <td style="border: solid #E3E3E3 1.0pt; border-right: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-left-alt: solid #E3E3E3 .75pt; mso-border-bottom-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center">&nbsp;</p>
    </td>
    <td style="border: solid #E3E3E3 1.0pt; border-left: none; mso-border-top-alt: solid #E3E3E3 .75pt; mso-border-bottom-alt: solid #E3E3E3 .75pt; mso-border-right-alt: solid #E3E3E3 .75pt; padding: 0cm 0cm 3.75pt 0cm;"><strong>Requested By</strong>: ##RequestedBy</td>
    </tr>
    </tbody>
    </table>
    </td>
    <td style="width: 5.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" width="5%">
    <p class="MsoNormal" style="line-height: 19.2pt;"><span style="font-size: 9.0pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
    </td>
    <td style="width: 40.0%; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top" width="40%">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: #FAFAFA; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes;">
    <td style="width: 100.0%; border: solid #E3E3E3 1.0pt; mso-border-alt: solid #E3E3E3 .75pt; padding: 7.5pt 0cm 1.5pt 3.75pt;" width="100%">
    <p style="text-align: center;" align="center"><span style="font-size: 10.5pt; color: #959595;">au2mator Request ID</span></p>
    <p style="text-align: center;" align="center"><u><span style="font-size: 12.0pt; color: black;"><a href="##RequestLink"><span style="color: black;">##REQUESTID</span></a></span></u></p>
    <p class="MsoNormal" style="text-align: center;" align="center"><span style="mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 2; box-sizing: border-box;">
    <td style="border: none; padding: 0cm 15.0pt 15.0pt 15.0pt; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm; box-sizing: border-box;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><strong><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">Dear ##UserDisplayname,</span></strong></p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 1; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="line-height: 19.2pt;"><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';">We finished the Request <strong>"##RequestTitle"</strong>!<br /> <br /> Here are the Result of the Request:<br /><b>##HTMLINFO&nbsp;</b><br /></span></p>
    <div>&nbsp;</div>
    <div>See the details of the Request</div>
    <div>##InputHTML</div>
    <div>&nbsp;</div>
    <div>&nbsp;</div>
    Kind regards,<br /> au2mator Self Service Team
    <p>&nbsp;</p>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 2; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="padding: 0cm 0cm 15.0pt 0cm; box-sizing: border-box;" valign="top">
    <p class="MsoNormal" style="text-align: center; line-height: 19.2pt;" align="center"><span style="font-size: 10.5pt; font-family: 'Helvetica',sans-serif; mso-fareast-font-family: 'Times New Roman';"><a style="border-radius: 3px; -webkit-border-radius: 3px; -moz-border-radius: 3px; display: inline-block;" href="##RequestLink"><strong><span style="color: white; border: solid #50D691 6.0pt; padding: 0cm; background: #50D691; text-decoration: none; text-underline: none;">View your Request</span></strong></a></span></p>
    </td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="mso-yfti-irow: 3; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="border: none; padding: 0cm 0cm 0cm 0cm; box-sizing: border-box;" valign="top">
    <table class="MsoNormalTable" style="width: 100.0%; mso-cellspacing: 0cm; background: #333333; mso-yfti-tbllook: 1184; mso-padding-alt: 0cm 0cm 0cm 0cm;" border="0" width="100%" cellspacing="0" cellpadding="0">
    <tbody>
    <tr style="mso-yfti-irow: 0; mso-yfti-firstrow: yes; mso-yfti-lastrow: yes; box-sizing: border-box;">
    <td style="width: 50.0%; border: none; border-right: solid lightgrey 1.0pt; mso-border-right-alt: solid lightgrey .75pt; padding: 22.5pt 15.0pt 22.5pt 15.0pt; box-sizing: border-box;" valign="top" width="50%">&nbsp;</td>
    <td style="width: 50.0%; padding: 22.5pt 15.0pt 22.5pt 15.0pt; box-sizing: border-box;" valign="top" width="50%">&nbsp;</td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </div>
    </td>
    <td style="padding: .75pt .75pt .75pt .75pt; box-sizing: border-box;" valign="top">&nbsp;</td>
    </tr>
    </tbody>
    </table>
    <p class="MsoNormal"><span style="mso-fareast-font-family: 'Times New Roman';">&nbsp;</span></p>
'@

    $html = $html.replace('##REQUESTID', $f_RequestID).replace('##UserDisplayname', $f_DisplaynameInitiatedBy).replace('##RequestTitle', $f_RequestTitle).replace('##EndDate', $f_EndDate).replace('##Status', $f_RequestStatus).replace('##RequestedBy', $f_InitiatedBy).replace('##HTMLINFO', $f_HTMLINFO).replace('##InputHTML', $f_InputHTML).replace('##RequestLink', $f_RequestLink)

    return $html
}

Function Send-ServiceMail ($HTMLBody, $ServiceName, $Recipient, $RequestID, $RequestStatus) {
    Write-au2matorLog -Type INFO -Text "Function Send-ServiceMail"
    $f_Subject = "au2mator - $ServiceName Request [$RequestID] - $RequestStatus"
    Write-au2matorLog -Type INFO -Text "Subject:  $f_Subject "
    Write-au2matorLog -Type INFO -Text "Recipient: $Recipient"

    try {
        if ($SMPTAuthentication) {

            if ($EnableSSLforSMTP) {
                Write-au2matorLog -Type INFO -Text "Run SMTP with Authentication and SSL"
                Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $SMTPcredential -UseSsl -Port $SMTPPort
            }
            else {
                Write-au2matorLog -Type INFO -Text "Run SMTP with Authentication and no SSL"
                Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Credential $SMTPcredential -Port $SMTPPort
            }
        }
        else {

            if ($EnableSSLforSMTP) {
                Write-au2matorLog -Type INFO -Text "Run SMTP without Authentication and SSL"
                Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -UseSsl -Port $SMTPPort
            }
            else {
                Write-au2matorLog -Type INFO -Text "Run SMTP without Authentication and no SSL"
                Send-MailMessage -SmtpServer $SMTPServer -To $Recipient -From $SMTPSender -Subject $f_Subject -Body $HTMLBody -BodyAsHtml -Priority high -Port $SMTPPort
            }
        }
    }
    catch {
        Write-au2matorLog -Type WARNING -Text "Error on sending Mail"
        Write-au2matorLog -Type WARNING -Text $Error
    }

}

Function Send-TeamsCard ($ServiceName, $Recipient, $RequestID, $RequestStatus, $au2matorReturn) {
    Write-au2matorLog -Type INFO -Text "Function Send Teams Card"

    switch ($RequestStatus) {
        "COMPLETED" { $statusColor = "Good" }
        "IN PROGRESS" { $statusColor = "Accent" }
        "ERROR" { $statusColor = "Attention" }
        "FAILED" { $statusColor = "Attention" }
        "WARNING" { $statusColor = "Warning" }
        Default { $statusColor = "Good" }
    }
    Write-au2matorLog -Type INFO -Text "Calculated the Statuscolor: $statusColor related to our Status: $RequestStatus"
    $URL = $PortalURL + "/Requests/SingleStatus?id=$RequestID"


    #Connect to GRAPH API
    $tokenBody = @{  
        Grant_Type = "password"  
        Scope      = "user.read%20openid%20profile%20offline_access"  
        Client_Id  = $clientId  
        username   = $User
        password   = $pw
        resource   = "https://graph.microsoft.com"
    }   

    $tokenResponse = Invoke-RestMethod "https://login.microsoftonline.com/common/oauth2/token" -Method Post -ContentType "application/x-www-form-urlencoded" -Body $tokenBody -ErrorAction STOP
    $headers = @{
        "Authorization" = "Bearer $($tokenResponse.access_token)"
        "Content-type"  = "application/json"
    }

    $GUID = (New-Guid).guid


    $GetRecipientUserURL = "https://graph.microsoft.com/v1.0/users/$Recipient"
    $RecipientUserID = (Invoke-RestMethod -Uri $GetRecipientUserURL -Method GET -Headers $headers).id
    $RecipientUserDisplayName = (Invoke-RestMethod -Uri $GetRecipientUserURL -Method GET -Headers $headers).displayname
    

    if ($ToChannel) {
        Write-au2matorLog -Type INFO -Text "Send Card to Channel"
        #Get ID for the Team
        $URLgetteamid = "https://graph.microsoft.com/v1.0/groups?$select=id,resourceProvisioningOptions"
        $TeamID = ((Invoke-RestMethod -Method GET -Uri $URLgetteamid  -Headers $headers).value | Where-Object -property displayName -value $TeamName -eq).id


        #Get ID for the Channel
        $URLgetchannelid = "https://graph.microsoft.com/v1.0/teams/$TeamID/channels"
        $ChannelID = ((Invoke-RestMethod -Method GET -Uri $URLgetchannelid  -Headers $headers).value | Where-Object -property displayName -value $ChannelName -eq).id

        #Send Message in channel

        $URLchatmessage = "https://graph.microsoft.com/v1.0/teams/$TeamID/channels/$ChannelID/messages"

    }

    if ($ToUser) {
        Write-au2matorLog -Type INFO -Text "Send Card to User"
        $GetBotUserURL = "https://graph.microsoft.com/v1.0/users/$User"

        $BotUserID = (Invoke-RestMethod -Uri $GetBotUserURL -Method GET  -Headers $headers).id
 
    
        $ChatBody = @"
    {
        "chatType": "oneOnOne",
        "members": [
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": "https://graph.microsoft.com/v1.0/users('$BotUserID')"
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": ["owner"],
            "user@odata.bind": "https://graph.microsoft.com/v1.0/users('$RecipientUserID')"
          }
        ]
      }
"@
        $URLchat = "https://graph.microsoft.com/v1.0/Chats"
    
        $ChatID = (Invoke-RestMethod -Uri $URLchat -Method POST -Body $ChatBody -Headers $headers).id
    
    
        $URLchatmessage = "https://graph.microsoft.com/v1.0/Chats/$ChatID/messages"

    }

   

    $BodyJsonTeam = @"
    {
        "subject": null,
        "body": {
            "contentType": "html",
            "content": "<attachment id=\"$GUID\"></attachment>"
        },
        
        "attachments": [
        {
          "id": "$GUID",
          "contentType": "application/vnd.microsoft.card.adaptive",
          "content": "{\n        \"type\": \"AdaptiveCard\",\n        \"$schema\": \"http://adaptivecards.io/schemas/adaptive-card.json\",\n        \"version\": \"1.2\",\n    \"speak\": \"<s>au2mator Request ID $RequestID is $RequestStatus</s><s>$au2matorReturn</s>\",\n    \"body\": [\n        {\n            \"type\": \"TextBlock\",\n            \"text\": \"Hi <at>$RecipientUserDisplayName</at>, \\nhere is the Result of your au2mator Service\",\n            \"wrap\": true,\n            \"w      eight\": \"Bolder\",\n            \"fontType\": \"Default\"\n        },{\n            \"type\": \"ColumnSet\",\n            \"columns\": [\n                {\n                    \"type\": \"Column\",\n                    \"width\": \"auto\",\n                    \"items\": [\n                        {\n                            \"type\": \"Image\",\n      \"url\": \"https://au2mator.com/wp-content/uploads/2018/02/HPLogoau2mator-1.png\",\n                            \"altText\": \"au2mator Logo\"\n                        }\n                    ]\n},\n                {\n                    \"type\": \"Column\",\n                    \"width\": \"stretch\",\n                    \"items\": [\n                        {\n                            \"type\": \"TextBlock\",\n                            \"text\": \"Status\",\n                            \"horizontalAlignment\": \"Right\",\n                            \"isSubtle\": true,\n                            \"wrap\": true\n                        },\n                        {\n                            \"type\": \"TextBlock\",\n                            \"text\": \"$RequestStatus\",\n                            \"horizontalAlignment\": \"Right\",\n                            \"spacing\": \"None\",\n                            \"size\": \"Large\",\n                            \"wrap\": true,\n                            \"color\": \"$statusColor\"\n}\n                    ]\n                }\n            ]\n        },\n        {\n            \"type\": \"ColumnSet\",\n            \"separator\": true,\n            \"spacing\": \"Medium\",\n            \"columns\": [\n                {\n                    \"type\": \"Column\",\n                    \"width\": \"stretch\",\n                    \"items\": [\n                        {\n                            \"type\": \"TextBlock\",\n                            \"text\": \"ServiceName\",\n                            \"wrap\": true\n                        },\n                        {\n                            \"type\": \"TextBlock\",\n     \"text\": \"RequestID\",\n                            \"wrap\": true\n                        }\n                    ]\n                },\n                {\n                    \"type\": \"Column\",\n                    \"width\": \"auto\",\n                    \"items\": [\n                        {\n                            \"type\": \"TextBlock\",\n                            \"horizontalAlignment\": \"Right\",\n                            \"isSubtle\": true,\n                            \"weight\": \"Bolder\",\n                            \"wrap\": true,\n                            \"text\": \"$ServiceName        \"\n                        },\n                        {\n                            \"type\": \"TextBlock\",\n                            \"text\": \"$RequestID\",\n                            \"horizontalAlignment\": \"Right\",\n                            \"spacing\": \"Small\",\n                            \"wrap\": true\n                        }\n                    ]\n                }\n            ]\n        },\n        {\n\"type\": \"ColumnSet\",\n            \"spacing\": \"Medium\",\n            \"separator\": true,\n            \"columns\": [\n                {\n                    \"type\": \"Column\",\n                    \"width\": 1,\n                    \"items\": [\n                        {\n                            \"type\": \"TextBlock\",\n                            \"text\": \"Request Return\",\n                            \"isSubtle\": true,\n                            \"weight\": \"Bolder\",\n                            \"wrap\": true\n                        }\n                    ]\n                }\n            ]\n        },\n        {\n\"type\": \"ColumnSet\",\n            \"spacing\": \"Small\",\n            \"columns\": [\n                {\n                    \"type\": \"Column\",\n                    \"width\": 1,\n                    \"items\": [\n                        {\n                            \"type\": \"TextBlock\",\n                            \"text\": \"$au2matorReturn \\n            \",\n\"isSubtle\": true,\n                            \"wrap\": true\n                        },\n                        {\n                            \"type\": \"ActionSet\",\n                            \"actions\": [\n                                {\n                                    \"type\": \"Action.OpenUrl\",\n                                    \"title\": \"Request Details\",\n                                    \"url\": \"$URL\",\n                                    \"style\": \"destructive\"\n                                }\n                            ],\n                            \"height\": \"stretch\"\n}\n                    ]\n                }\n            ]\n        }\n    ],\n     \"msteams\": {\n      \"entities\": [\n        {\n          \"type\": \"mention\",\n          \"text\": \"\u003cat\u003e$RecipientUserDisplayName\u003c/at\u003e\",\n          \"mentioned\": {\n            \"id\": \"8:orgid:$RecipientUserID\",\n            \"name\": \"$RecipientUserDisplayName\"\n          }\n        }\n      ]\n    }\n        }"
        }
      ]
    }
"@


    Write-au2matorLog -Type INFO -Text "Send Card"
    Invoke-RestMethod -Method POST -Uri $URLchatmessage -Body $BodyJsonTeam -Headers $headers
    
}


#endregion Functions


#region CustomFunctions

#
#
#
#


#endregion CustomFunctions

#region Script
Write-au2matorLog -Type INFO -Text "Start Script"


if ($DoImportPSSession) {

    Write-au2matorLog -Type INFO -Text "Import-Pssession"
    $PSSession = New-PSSession -ComputerName $PSRemotingServer
    Import-PSSession -Session $PSSession -DisableNameChecking -AllowClobber 
}

#Check for Modules if installed
Write-au2matorLog -Type INFO -Text "Try to install all PowerShell Modules"
foreach ($Module in $Modules) {
    if (Get-Module -ListAvailable -Name $Module) {
        Write-au2matorLog -Type INFO -Text "Module is already installed:  $Module"        
    }
    else {
        Write-au2matorLog -Type INFO -Text "Module is not installed, try simple method:  $Module"
        try {

            Install-Module $Module -Force -Confirm:$false
            Write-au2matorLog -Type INFO -Text "Module was installed the simple way:  $Module"

        }
        catch {
            Write-au2matorLog -Type INFO -Text "Module is not installed, try the advanced way:  $Module"
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Install-PackageProvider -Name NuGet  -MinimumVersion 2.8.5.201 -Force
                Install-Module $Module -Force -Confirm:$false
                Write-au2matorLog -Type INFO -Text "Module was installed the advanced way:  $Module"

            }
            catch {
                Write-au2matorLog -Type ERROR -Text "could not install module:  $Module"
                $au2matorReturn = "could not install module:  $Module, Error: $Error"
                $AdditionalHTML = "could not install module:  $Module, Error: $Error
                "
                $Status = "ERROR"
            }
        }
    }
    Write-au2matorLog -Type INFO -Text "Import Module:  $Module"
    Import-module $Module
}

#region CustomCode
Write-au2matorLog -Type INFO -Text "Start Custom Code"
$error.Clear()



try {
    Write-au2matorLog -Type INFO -Text "Try to set TTL"
    if ($c_Type -eq "Duration") {
           
        Write-au2matorLog -Type INFO -Text "TTL is calculated with Duration"
        $ttl = New-TimeSpan -Minutes $c_TTL
    }
        
    if ($c_Type -eq "EndDate") {
        Write-au2matorLog -Type INFO -Text "TTL is calculated with EndDate"

        $Now = ([datetime]::ParseExact((Get-Date -Format "dd-MM-yy hh:mm:ss" ), "dd-MM-yy hh:mm:ss", [cultureinfo]::InvariantCulture))
        $End = ([datetime]::ParseExact($c_EndDate, "dd-MM-yy hh:mm:ss", [cultureinfo]::InvariantCulture))
        Write-au2matorLog -Type INFO -Text "Now: $Now"
        Write-au2matorLog -Type INFO -Text "Selected: $c_EndDate"
        Write-au2matorLog -Type INFO -Text "End: $End"
            
        $ttl = New-TimeSpan -End $End
        $ttl = New-TimeSpan -Minutes $ttl.TotalMinutes
    }

    Write-au2matorLog -Type INFO -Text "The calculated TTL is: $($ttl.Days) (Days) $($ttl.Hours) (Hours) $($ttl.Minutes) (Minutes) "

    $ADUser = Get-ADUser -Identity $c_User
    $ADGroup = Get-ADGroup -Identity $c_Group
        
    try {
        Write-au2matorLog -Type INFO -Text "Try to add user to Group with TTL"
        Add-ADGroupMember -Identity $ADGroup.SamAccountName -Members  $ADUser.SamAccountName -MemberTimeToLive $ttl


        $au2matorReturn = "The User $($ADUser.Name) was added to the Group $($ADGroup.Name) with a limit of  $($ttl.Days) (Days) $($ttl.Hours) (Hours) $($ttl.Minutes) (Minutes) "
        $TeamsReturn="The User $($ADUser.Name) was added to the Group $($ADGroup.Name) with a limit of  $($ttl.Days) (Days) $($ttl.Hours) (Hours) $($ttl.Minutes) (Minutes) "
        $AdditionalHTML = "The User $($ADUser.Name) was added to the Group $($ADGroup.Name) with a limit of  <br>$($ttl.Days) (Days) <br>$($ttl.Hours) (Hours) <br>$($ttl.Minutes) (Minutes) 
        <br>
        "
        $Status = "COMPLETED"
        
    }
    catch {
        Write-au2matorLog -Type ERROR -Text "Could not add User to Group"
        Write-au2matorLog -Type ERROR -Text "Error: $Error"

        $au2matorReturn = "Could not add User to Group, Error: $Error"
        $TeamsReturn="Could not add User to Group"
        $AdditionalHTML = "Could not add User to Group
        <br>
        Error: $Error
            "
        $Status = "ERROR"
    }
}
catch {
    Write-au2matorLog -Type ERROR -Text "Error on getting the TTL, Error: $Error"
    Write-au2matorLog -Type ERROR -Text "Error on getting the TTL: $Error"

    $au2matorReturn = "Error on getting the TTL, Error: $Error"
    $TeamsReturn="Error on getting the TTL"
    $AdditionalHTML = "Error on getting the TTL
        <br>
        Error: $Error
            "
    $Status = "ERROR"
}



#endregion CustomCode
#endregion Script

#region Return


Write-au2matorLog -Type INFO -Text "Service finished"

if ($SendMailToInitiatedByUser -or $SendMailToTargetUser -or $SendTeamsCardToInitiatedByUser -or $SendTeamsCardToTargetUser) {
    Write-au2matorLog -Type INFO -Text "We need to Send Mail or Teams Chat, so get all Infos"
    $UserInput = Get-UserInput -RequestID $RequestId
    $HTML = Get-MailContent -RequestID $RequestId -RequestTitle $Service -EndDate $UserInput.ApprovedTime -TargetUserId $UserInput.TargetUserId -InitiatedBy $UserInput.InitiatedBy -Status $Status -PortalURL $PortalURL  -AdditionalHTML $AdditionalHTML -InputHTML $UserInput.html
}

if ($SendMailToInitiatedByUser) {
    Write-au2matorLog -Type INFO -Text "Send Mail to Initiated By User"
    Send-ServiceMail -HTMLBody $HTML -RequestID $RequestId -Recipient $($UserInput.MailInitiatedBy) -RequestStatus $Status -ServiceName $Service
}

if ($SendMailToTargetUser) {
    Write-au2matorLog -Type INFO -Text "Send Mail to Target User"
    Send-ServiceMail -HTMLBody $HTML -RequestID $RequestId -Recipient $($UserInput.MailTarget) -RequestStatus $Status -ServiceName $Service
}


if ($SendTeamsCardToInitiatedByUser) {
    Write-au2matorLog -Type INFO -Text "Send Teams Card to InitiatedBy"
    Send-TeamsCard  -RequestID $RequestId -Recipient $($UserInput.MailInitiatedBy) -RequestStatus $Status -ServiceName $Service -au2matorReturn $TeamsReturn
}

if ($SendTeamsCardToTargetUser) {
    Write-au2matorLog -Type INFO -Text "Send Teams Card to Target User"
    Send-TeamsCard  -RequestID $RequestId -Recipient $($UserInput.MailTarget) -RequestStatus $Status -ServiceName $Service -au2matorReturn $TeamsReturn
}

return $au2matorReturn
#endregion Return
