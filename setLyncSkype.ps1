<#
  Powershell script to set both Lync and Skype status to either Online or Away

  Usage: C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe %HOMEPATH%\bin\setLyncSkype\setLyncSkype.ps1 Online|Away

  Requires:
    Lync SDK
      Lync 2010 SDK (16MB)   http://www.microsoft.com/en-us/download/details.aspx?id=18898
    Skype4COM
      Download Skype4COM version 1.0.38 (zip archive) 1.8MB,  http://dev.skype.com/developer/resources/Skype4COM-1.0.38.0.zip
      Unpack zip file somewhere (make dir Skype4COM under normal skype install dir): C:\Program Files (x86)\Skype\Skype4COM
      CMD as Admin: regsvr32 "C:\Program Files (x86)\Skype\Skype4COM\Skype4COM.dll"
      You should get a response that it was registered successfully: DllRegisterServer ... succeeded.

  Resources:
    http://www.ravichaganti.com/blog/?p=2613
    http://blogs.technet.com/b/csps/archive/2011/05/05/sendim.aspx?Redirected=true
    http://stackoverflow.com/questions/13495820/simple-example-tutorial-showing-how-to-send-lync-messages-from-a-console-app?rq=1
    http://www.wroolie.co.uk/2009/09/automate-skype-status-with-powershell/
    http://dmitrysotnikov.wordpress.com/2012/11/09/powershell-script-to-set-skype-status-text-to-latest-blog-or-twitter-update/
    https://support.skype.com/en/faq/FA171/can-i-run-skype-for-windows-desktop-from-the-command-line

  Orig Author/Date/(c): awcoleman/20130325/awcoleman    License: MIT (http://opensource.org/licenses/mit-license.html)
  Last Update: awcoleman/20130325
#>

if ( $args.count -ne 1 ) {
	Write-Host "`nStatus argument is missing. Exiting.`n"
	Exit
}
$inputStatus = $args[0]

if ( $inputStatus.ToLower() -eq "online" ) {
	$inputLyncStatus = "available"
	$inputLyncStatusNumber = 3500
	$inputSkypeStatus = "ONLINE"
} elseif ( $inputStatus.ToLower() -eq "away" ) {
	$inputLyncStatus = "be-right-back"
	$inputLyncStatusNumber = 15500
	$inputSkypeStatus = "AWAY"
} else {
	Write-Host "`nStatus argument is unknown. Exiting.`n"
	Exit
}

#Lync section
import-module "C:\Program Files (x86)\Microsoft Lync\SDK\Assemblies\Desktop\Microsoft.Lync.Model.Dll"
$lyncClient = [Microsoft.Lync.Model.LyncClient]::GetClient()
$lync = $lyncClient.Self
$lyncContactInfo = new-object 'System.Collections.Generic.Dictionary[Microsoft.Lync.Model.PublishableContactInformationType, object]'
$lyncContactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::Availability,$inputLyncStatusNumber)
$lyncContactInfo.Add([Microsoft.Lync.Model.PublishableContactInformationType]::ActivityId,$inputLyncStatus)
$lyncSetReturn = $lync.BeginPublishContactInformation($lyncContactInfo, $null, $null)

#Skype section
$skype = New-Object -COM "Skype4COM.Skype"
$skypeInputOnlineStatus = $skype.Convert.TextToUserStatus($inputSkypeStatus)
$skype.ChangeUserStatus($skypeInputOnlineStatus)

#Finished
Exit