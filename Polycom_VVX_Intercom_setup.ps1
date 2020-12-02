#########################################################
# Authors:     Ryan Vanderhoff
# Created:     2018-11-06
# Description: Change Intercom setup on Polycom Phones 
# NOTES
#  Rvanderhoff:12-02-2020 - Works with Polycom 311 and 411 phones 
#   
#########################################################

#Checks IE Setting and updates
$registryPath = "HKCU:\Software\Microsoft\Internet Explorer\Main"
$Name = "Isolation"
$value = "PMEM"
$ie_settings = Get-ItemProperty -Path $registryPath -Name $Name
If($ie_settings.isolation -eq "PMIL") {
    Set-ItemProperty -Path $registryPath -Name $name -Value $value
    Write-Host "Enabled Enhanced Protected Mode in IE"
}

#Enter Username and Password of Polycom admin user below
$username = "Polycom"
$password = "456" | ConvertTo-SecureString -AsPlainText -Force
#Stores credential securely
$cred = New-Object System.Management.Automation.PSCredential($username,$password)

add-type @"
        using System.Net;
        using System.Security.Cryptography.X509Certificates;

            public class IDontCarePolicy : ICertificatePolicy {
            public IDontCarePolicy() {}
            public bool CheckValidationResult(
                ServicePoint sPoint, X509Certificate cert,
                WebRequest wRequest, int certProb) {
                return true;
            }
        }
"@
[System.Net.ServicePointManager]::CertificatePolicy = new-object IDontCarePolicy 
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls11

#Set Prefix range here
$ip_prefix = "192.168.1."
$IP_Range = @()

#Get IP's for VOIP Phones and place in array. Change starting and ending range as needed. First $i value is starting. 
#Number after -lt is 1 above the last ip checked.      
for($i=2 ;$i -lt 255; $i++) {
    $ip_new = $ip_prefix + $i
    $IP_Range += $ip_new
}

#Collect devices that are online and store into an good array
Write-host Collecting phones that are online.....
$ping = New-Object System.Net.Networkinformation.ping

$ip_good = @()
foreach($ip in $IP_Range) {
    #$test1 = Test-Connection -ComputerName $ip -Count 1 -Quiet -BufferSize 16
    $connection = $ping.Send("$ip",200)
    Write-host Checking IP $ip
    if($connection.Status -eq "Success") {
        $ip_good += $ip
    }
}

##Check if known phone type. 
#Create arrays for each phone type. Add additional arrays for any additional phone types
$Polycom = @()
$Yealink = @()
$unknown = @()

#Loop through known online IP's and checks web page for phone type using IE
foreach($ip in $ip_good) {
    $ie = new-object -ComObject "InternetExplorer.Application"
    $requesturi = "https://$($ip)"
    #$ie.visible = $true
    $ie.silent = $true
    $ie.navigate($requestUri)
    while($ie.Readystate -eq 4) { Start-Sleep -Milliseconds 100 }
    Start-Sleep -s 10
    if ($ie.document.url -Match "invalidcert") {
        "Bypassing SSL Certificate Error Page for $requestUri"
        $sslbypass=$ie.Document.getElementsByTagName("a") | where-object {$_.id -eq "overridelink"}
        $sslbypass.click()
        "Waiting for 20 seconds while final page loads"
        start-sleep -s 20
    }
    $doc = $ie.Document
    if($doc.title -match "Polycom") {
        Write-Host "Polycom"
        $Polycom += $ip
    } Elseif($doc.title -match "Yealink") {
        Write-Host "Yealink"
        $Yealink += $ip
    } Else {
        Write-host "Unknown"
        $Unknown += $ip
    }
    $ie.Quit()
    start-sleep -s 5
}

#Arrays for if RestAPI is enabled
$deviceinfo2 = @()
$EnableRestAPI = @()

##Change the config on each Polycom phone
#Loops through each Polycom detected phone
foreach ($phone_ip in $Polycom) {
    #Gets Device info from Rest API
    Write-Host "Checking $phone_ip Polycom config"
    $deviceinfo = Invoke-RestMethod -Uri "https://$phone_ip/api/v1/mgmt/device/info" -Credential $cred  -Method GET -ContentType "application/json"  -TimeoutSec 2
    if($deviceinfo) {
        $deviceinfo2 += $deviceinfo.data
    } Else {
        $EnableRestAPI += $phone_ip
    }
    $deviceinfo = ""
}

##Enable RestAPI if necessary
#Goes through all phones until Rest API is enabled on all phones
while($EnableRestAPI.count -gt 0) {
    #Loops through each Polycom phone that needs the Rest API enabled
    foreach($phone_ip in $EnableRestAPI) {
        Write-Host "Enabling Rest API for $phone_ip"
        $ie = new-object -ComObject "InternetExplorer.Application"
        $requesturi = "https://$($phone_ip)"
        #Uncomment below line if you need to see the Web browser in action to debug with and comment $ie.silent = $true 
        #$ie.visible = $true
        $ie.silent = $true
        $ie.navigate($requestUri)
        while($ie.ReadyState -ne 4) {start-sleep -m 100}
        Start-Sleep -s 5
        if ($ie.document.url -Match "invalidcert"){
            Write-Host "Bypassing SSL Certificate Error Page for $requesturi"
            $sslbypass=$ie.Document.getElementsByTagName("a") | where-object {$_.id -eq "overridelink"}
            $sslbypass.click()
            Write-Host "Waiting for 10 seconds while final page loads"
            start-sleep 10
        }
        $login_check = $ie.Document.body.getElementsByTagName("th") | Where {$_.outertext -match "Enter Login Information"}
        Start-Sleep -s 5
        if($login_check.outerText -eq "Enter Login Information") {
        start-sleep -s 2
        $pass = $ie.Document.body.getElementsByTagName("input") | Where {$_.name -eq "password"}
        $pass.value = $cred.GetNetworkCredential().password
        start-sleep -s 2
        $submit = $ie.Document.body.getElementsByTagName("input") | Where {$_.value -eq "Submit"}
        $submit.Click()
        start-sleep -s 2
        }
        $doc = $ie.Document
        start-sleep -s 5
        $settings = $doc.body.getElementsByTagName('a') | Where{$_.innerhtml -match "Applications"}
        $settings.click()
        start-sleep -s 3
        $rest = $doc.body.getElementsByTagName('input') | Where {$_.name -eq "11" -and $_.type -eq "radio"}
        $rest[0].checked = "true"
        start-sleep -s 2
        $save = $doc.body.getElementsByTagName("button") | Where{$_.textcontent -match "Save"}
        $save.click()
        Start-Sleep -s 3
        $yes = $doc.body.getElementsByTagName("button") | Where{$_.textcontent -match "yes"}
        $yes.click()
        start-sleep -s 5
        $logout = $doc.body.getElementsByTagName('a') | Where{$_.innertext -match "Log Out"}
        $logout.click()
        Start-Sleep -s 3
        $ie.Quit()
        start-sleep -s 5
    }
    #run Polycom API scan again
    $deviceinfo2 = @()
    $EnableRestAPI = @()
    foreach ($phone_ip in $Polycom) {
        #Gets Device info
        Write-Host "Checking $phone_ip Polycom config"
        $deviceinfo = Invoke-RestMethod -Uri "https://$phone_ip/api/v1/mgmt/device/info" -Credential $cred  -Method GET -ContentType "application/json"  -TimeoutSec 2
        if($deviceinfo) {
        $deviceinfo2 += $deviceinfo.data
    } Else {
        $EnableRestAPI += $phone_ip
    }
    $deviceinfo = ""
    }
}

#Check if Intercom setup if needed
$intercom = @()
foreach ($device in $deviceinfo2) {
    if ($device) {
        $firmware = $device.FirmwareRelease
        $model = $device.ModelNumber
        $mac = $device.MACAddress
        $ip = $device.IPV4Address
        $lineinfo = Invoke-RestMethod -Uri "https://$ip/api/v1/mgmt/lineInfo" -Credential $cred  -Method GET -ContentType "application/json"  -TimeoutSec 2
        $linedata = $lineinfo.data | Where {$_.LineNumber -eq 1}
        $lines = $lineinfo.data.Count
        $Ext = $linedata.Label
        $number = $linedata.SIPAddress

        ##Get intercom setting
        #Create JSON message
        $GetConfig_JSON = @"
{
    `"data`":
    [
    `"voIpProt.SIP.intercom.alertInfo`",
    `"se.rt.ringAutoAnswer.ringer`",
    `"voIpProt.SIP.alertInfo.1.class`"
    ]
}
"@
    
        #Sends REST API call to check current setting
        $get_config = Invoke-RestMethod -Uri "https://$ip/api/v1/mgmt/config/get" -Credential $cred  -Method POST -ContentType "application/json" -Body $GetConfig_JSON -TimeoutSec 2 

        #Checks if ringAutoAnswer is enabled for Intercom
        if($get_config.data.'voIpProt.SIP.alertInfo.1.class'.Value -eq "ringAutoAnswer") {
             
        } Else {
            $intercom += $ip
        }
    }
}

#Change the config on each phone if needed
foreach ($phone_ip in $intercom) {
    #Remove application button from homescreen, set intercom to RingAutoAnswer, and change ringer of RingerAutoAnswer
    $SendConfig_JSON = @"
    {
        `"data`":
        {
        `"voIpProt.SIP.intercom.alertInfo`": `"unused;info=alert-autoanswer;delay=0`",
        `"voIpProt.SIP.alertInfo.1.class`": `"ringAutoAnswer`",
        `"se.rt.ringAutoAnswer.ringer`": `"ringer10`",
        `"homeScreen.application.enable": "0"
        }
    }
"@

    #send config to phone -- Phone will reboot ;REST API must be enabled prior to using
    Write-host Updating $phone_ip
    $send_config = Invoke-RestMethod -Uri "https://$phone_ip/api/v1/mgmt/config/set" -Credential $cred  -Method Post -ContentType "application/json"  -Body $SendConfig_JSON -TimeoutSec 2
}