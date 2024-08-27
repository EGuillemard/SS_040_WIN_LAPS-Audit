er# =======================================================
# NAME: SS_040_WIN_LAPS-Audit.ps1
# AUTHOR: GUILLEMARD, Erwan, PERSONNAL PROPRIETY
# DATE: 2024/08/25
#
# KEYWORDS: WINDOWS, LAPS, AUDIT
# VERSION 1.0
# 2024/08/25 - 1.0 : Script creation
# COMMENTS: 
#
#Requires -Version 2.0
# =======================================================


#Global Constants
#
#Format current date and cast one to DateTime
$StringDate=Get-Date -Format "dd/MM/yyyy hh:mm:ss"
$CurrentTime=[DateTime]::ParseExact($StringDate, "dd/MM/yyyy hh:mm:ss", $null)

#Constants dedied to HTML Report function
$_const_background_color="#BDCB04"
$_const_frontpolicy_color="#26313B"
$_const_body_color="#72797F"
$_const_author="Erwan GUILLEMARD"
$_const_version="1.0"
$_const_date="2024-08-25"

#Constants SMTP
$From="srvAD@contoso.lan"
$To="johnsmith@gmail.com"
$SMTPServer="srv.contoso.lan"
$Port=25

function getADComputersObject {
<#
    .SYNOPSIS
        Function to get Computer object from DC
    .DESCRIPTION
        Return DCs, Servers or computers from Domain if types are specified
    .PARAMETER
        String - role..................: Specify type of computer object than which return
    .EXAMPLE
        getADComputersObject -role $role
    .INPUTS
        Type:[String]
    .OUTPUTS
        Return ObjectArray
    .NOTES
        NAME:  Creation de la fonction
        AUTHOR: GUILLEMARD Erwan
        LASTEDIT: 2024/08/25
        VERSION:1.0.0 Function Create 
    .LINK
#>
    param(
        [Parameter(Mandatory=$true)][String]$role
    )
    $listObjectComputer=@()
    switch ($role) {
        #Get only computer exept DC and Servers from computersObjects
        "Computers" {
            $listObjectComputer=Get-ADComputer -Filter 'OperatingSystem -notlike "*Windows Server*"' -Properties OperatingSystem | Where-Object {$_.DistinguishedName -notlike "*Domain Controllers*"} | Select SamAccountName,Enabled,OperatingSystem
            break
        }
        #Get only servers exept DC from computers Objects
        "Servers" {
            $listObjectComputer=Get-ADComputer -Filter 'OperatingSystem -like "*Windows Server*"' -Properties OperatingSystem | Where-Object {$_.DistinguishedName -notlike "*Domain Controllers*"} | Select SamAccountName,Enabled,OperatingSystem
            break
        }
        #Get only DC exept Servers from computers Objects. SAM and statusEnabled Porperties only
        "DCs" {
            $listObjectComputer=Get-ADComputer -Filter 'OperatingSystem -like "*Windows Server*"' -Properties OperatingSystem | Where-Object {$_.DistinguishedName -like "*Domain Controllers*"} | Select-Object SamAccountName,Enabled,OperatingSystem
            break
        }
        default {
            $null
            break
        }
    }
    return $listObjectComputer
}

function getLAPSAudit {
<#
    .SYNOPSIS
        Function to get LAPS information from object computers (OnPremise only. AzureAD is not currently supported)
    .DESCRIPTION
        Return LAPS informations (enabled or not), next rotation date or if the renewal date is obsolete
    .PARAMETER
        Object[] - CIs..................: Specify the computer object array which are join and enable at the OnPremise domain
    .EXAMPLE
        getAPSAudit -CIs $CIs
    .INPUTS
        Type:[Object[]]
    .OUTPUTS
        Return ObjectArray custom with role (DC, Server or Computer) and all items associated
    .NOTES
        NAME:  Creation de la fonction
        AUTHOR: GUILLEMARD Erwan
        LASTEDIT: 2024/08/25
        VERSION:1.0.0 Function Create 
    .LINK
#>
    param(
        [Parameter(Mandatory=$true)][Object[]]$CIs
    )
    $listObjLAPSCI=@()
    foreach($ci in $CIs){
        $lapsInformation=Get-LapsADPassword -Identity $ci.SamAccountName
        #LAPS enabled on the CI
        if($lapsInformation.ExpirationTimestamp){
            $lapsEnabled=$true
            $expirationDate=Get-Date ($lapsInformation.ExpirationTimestamp) -Format "dd/MM/yyyy hh:mm:ss"
            #Password will be renewal
            if($CurrentTime -lt $lapsInformation.ExpirationTimestamp){
                $daysOngoing= New-TimeSpan -Start $CurrentTime -End $expirationDate
                $lapsRenewalStatus="To Be Renewal"
            }
            #Password hasn't be renewal
            else{
                $daysOngoing= New-TimeSpan -Start $CurrentTime -End $expirationDate
                $lapsRenewalStatus="Renewal Exceed"
            }
        }
        #LAPS NOT enabled on the CI
        else{
            $lapsEnabled=$false
            $daysOngoing=""
            $expirationDate=""
            $lapsRenewalStatus=""
        }
        $lapsItem=[PSCustomObject]@{
            CI_NAME=$ci.SamAccountName
            CI_STATUS=$ci.Enabled
            CI_SE=$ci.OperatingSystem
            LAPS_STATUS=$lapsEnabled
            LAPS_DAYSRENEW=$daysOngoing.Days
            LAPS_NEXTEXPIRATIONDATE=$expirationDate
            LAPS_ACTION=$lapsRenewalStatus
        }
        $listObjLAPSCI+=$lapsItem
    }
    return $listObjLAPSCI
}

function getLAPSHTML_Report {
<#
    .SYNOPSIS
        Function to generated the HTML report with LAPS informations
    .DESCRIPTION
        Return and build an HTML report with LAPS informations with isolated DC, Servers and computers on three differents parts 
    .PARAMETER
        Object[] - arrayComputerObjects..................: Specify the computer object array which are join and enable at the OnPremise domain
    .EXAMPLE
        getLAPSHTML_Report -arrayComputerObjects $arrayComputerObjects
    .INPUTS
        Type:[Object[]]
    .OUTPUTS
        Return String
    .NOTES
        NAME:  Creation de la fonction
        AUTHOR: GUILLEMARD Erwan
        LASTEDIT: 2024/08/25
        VERSION:1.0.0 Function Create 
    .LINK
#>
    param(
        [Parameter(Mandatory=$true)][Object[]]$arrayComputerObjects
    )
    $body="<html>"
    $body+="<head>"
    $body+="<meta http-equiv=`"Content-Type`" content=`"text/html`"; charset=UTF-8`" />"
    $body+="<meta http-equiv=`"X-UA-Compatible`" content=`"IE=EmulateIE7`" />"
    $body+="<style>"
	$body+="table, th, td {border: 1px solid black; border-collapse: collapse;}"
    $body+="</style>"
    $body+="</head>"

    $body+="<body>"
    $body+="<table>"
	$body+="<thead>"
	$body+="<tr>"
	$body+='<th colspan="2">'
	$body+="<center>"
	$body+='<h1 style="background-color:$($_const_background_color); color:$($_const_frontpolicy_color)";>LAPS Audit generated automatically</h1>'
	$body+="</center>"
	$body+="</th>"
	$body+="</thead>"

    $body+="<tbody>"
	$body+="<tr>"
	$body+="<td>"
	$body+="<h2 style=`"color:$($_const_frontpolicy_color)`";>SS_040_WIN_LAPS-Audit</h2>"
	$body+="</td>"
	$body+="<td>"
	$body+="<p align=`"right`" style=`"color:$($_const_frontpolicy_color)`";>$($_const_author)<b><u> :Author</u></b><br/>$($_const_version)<b><u> :Version</u></b><br/>$($_const_date)<b><u> :Date</u></b></p>"
	$body+="</td>"
	$body+="</tr>"
    
    $body+="</tbody>"
    $body+="</table>"
    $body+="<br/>"

    foreach ($computerObjects in $arrayComputerObjects) {
        switch ($computerObjects.OBJ_ROLE) {
            "Computers" {
                $title="Computers"      
            break
            }
            "DCs" {
                $title="Domain Controllers"    
            break
            }
            "Servers" {
                $title="Servers"
            break
            }
            default {
                $title="Unknown Items"
            break
            }
        }
        $body+="<tr>"
        $body+="<h3 style=`"color:$($_const_frontpolicy_color)`";>$($title)</h3>"
        $body+="</tr>"
        $body+="<tr>"
        $body+="<table>"
        $body+="<thead>"
        $body+="<tr>"
        $body+="<th>Hostname</th>"
        $body+="<th>Status</th>"
        $body+="<th>OS Type</th>"
        $body+="<th>LAPS Status</th>"
        $body+="<th>Renewed Date</th>"
        $body+="<th>Days before next rotation</th>"
        $body+="<th>LAPS Actions</th>"
        $body+="</tr>"
        $body+="</thead>"
        $body+="<tbody>"
        foreach ($itemComputer in $computerObjects.OBJ_CUSTOM_DATA){
            #Check LAPS status to define color front
            if($itemComputer.LAPS_STATUS){
                $colorStatus="style=`"color:#03B000`""
            }
            else{
                $colorStatus="style=`"color:#F83838`""
            }
            #Check LAPS Action to define color front
            if($itemComputer.LAPS_ACTION -like "To Be Renewal"){
                $colorActions="style=`"color:#03B000`""
            }
            else{
                $colorActions="style=`"color:#F83838`""
            }
            $body+="<tr>"
            $body+="<td>$($itemComputer.CI_NAME)</td>"
            $body+="<td>$($itemComputer.CI_STATUS)</td>"
            $body+="<td>$($itemComputer.CI_SE)</td>"
            $body+="<td $($colorStatus)>$($itemComputer.LAPS_STATUS)</td>"
            $body+="<td>$($itemComputer.LAPS_NEXTEXPIRATIONDATE)</td>"
            $body+="<td>$($itemComputer.LAPS_DAYSRENEW)</td>"
            $body+="<td $($colorActions)>$($itemComputer.LAPS_ACTION)</td>"
            $body+="</tr>"
        }
        $body+="</tbody>"
        $body+="</table>"
        $body+="</tr>"

    }

	$body+="</tbody>"
    $body+="</table>"
    $body+="<i>"
	$body+='<p style="color:red;">For all emergencies, scripts assistances please contact Erwan GUILLEMARD</p>'
	$body+="</i>"
    $body+="</body>"
    $body+="</html>"
    $body > C:\Users\Administrateur\Desktop\test.html
    return $body
}

function sendSMTPMail {
<#
    .SYNOPSIS
        Function to send mail across SMTP relay
    .DESCRIPTION
        Send a mail to an SMTP server deployed on a LAN server to relaying mail to SMTP over SSL or TLS (tcp:465 or tcp:587)
         _____________________________________________________________________________________________
        |                 tcp:25                                    tcp:465/tcp:587
        |Apps or Server ----------> Internal SMTP Server (relay) ----------------------> SMTP Server
        |_____________________________________________________________________________________________
    .PARAMETER
        String - subject..................: Specify the mail subject
        String - htmlBody.................: Specify the mail body as HTLM format
    .EXAMPLE
        sendSMTPMail -subject $subject -htmlBody $htmlBody
    .INPUTS
        Type:[String]
        Type:[String]
    .OUTPUTS
        N/A
    .NOTES
        NAME:  Creation de la fonction
        AUTHOR: GUILLEMARD Erwan
        LASTEDIT: 2024/08/27
        VERSION:1.0.0 Function Create 
    .LINK
#>
    param(
        [Parameter(Mandatory=$true)][String]$subject,
        [Parameter(Mandatory=$true)][String]$htmlBody
    )
    Send-MailMessage -From $From `
                        -To $To `
                        -Subject $Subject `
                        -BodyAsHtml $htmlBody `
                        -SmtpServer $SMTPServer `
                        -Port $Port
}

$listComputerObjects=@()
$dcs=getADComputersObject -role "DCs"
$result=getLAPSAudit -CIs $dcs
$objectComputer=[PSCustomObject]@{
    OBJ_ROLE="DCs"
    OBJ_CUSTOM_DATA=$result
}
$listComputerObjects+=$objectComputer
$servers=getADComputersObject -role "Servers"
$result=getLAPSAudit -CIs $servers
$objectComputer=[PSCustomObject]@{
    OBJ_ROLE="Servers"
    OBJ_CUSTOM_DATA=$result
}
$listComputerObjects+=$objectComputer
$computers=getADComputersObject -role "Computers"
$result=getLAPSAudit -CIs $computers
$objectComputer=[PSCustomObject]@{
    OBJ_ROLE="Computers"
    OBJ_CUSTOM_DATA=$result
}
$listComputerObjects+=$objectComputer
$htmlBody=getLAPSHTML_Report -arrayComputerObjects $listComputerObjects
sendSMTPMail -Subject "LAPS Audit" `
             -htmlBody $htmlBody
