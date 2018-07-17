<# SSWTemplateScript - 
 #
 #  Version     Author          Date            Comment  
 #  1.0         Greg Harris     12/03/2018      Initial Version - Based on SSWLoginScript.bat
 #  1.1         Kaique Biancatti07/06/2018      Added the correct link to GitHub and added TLS options to connect to HTTPS. Also added name prompt.
 #  1.2         Kaique Biancatti08/06/2018      Added self elevation of PowerShell script, comments, backup logic, and reorganizing of code.
 #  1.3	        Kaique Biancatti16/07/2018      Added open notepad with log at the end of script.
 #  1.4         Kaique Biancatti17/07/2018      Added time sync with Sydney server.
 #>

param (    
    [string]$username = ''
)

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())

If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -eq $False) {
    #Write-Host 'Script actions not performed. This script MUST be run as an Administrator.' -ForegroundColor Red
    #exit
#}
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File",('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}

#This line sets the variable with the current GitHub project with all our Templates, and creates our LogFile.
Set-Variable -Name 'ScriptTemplateSource' -Value 'https://github.com/SSWConsulting/LoginScript/raw/master/'
Set-Variable -Name 'ScriptLogFile' -Value 'C:\SSWTemplateScript_LastRun.log'

Set-Content -Path $ScriptLogFile -Value 'SSWTemplateScript log' -Force

Write-Host 'This PowerShell script copies SSW Template Files from ' $ScriptTemplateSource ' to your %AppData%\Microsoft\Templates\ folder'
Write-Host 'Please make sure that Word, Powerpoint and Outlook are closed. Open templates will not be replaced'

#This sets the security protocol to use all TLS versions. Without this, Powershell will use TLS1.0 which GitHub does not accept.
[Net.ServicePointManager]::SecurityProtocol =  "tls12, tls11, tls"

# This bit tries to find if the user is running on a domain account.
$domain = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[0]
if ($domain -eq 'SSW2000') {
    if($username -eq '') {
    $username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[1]
    }
    Add-Content -Path $ScriptLogFile -Value 'Domain username found'
}

if ($username -eq '') {
    Write-Host 'Username parameter required if not run on a SSW2000 domain account. Please input username on the pop-up box.' -ForegroundColor Red

	#Calling a VB Prompt for the user if there is no username set    
	[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
	$username = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your username with correct capitals and no domain - Eg. AdamCogan and not SSW2000\adamcogan or adamcogan@ssw.com.au", "Please input your SSW username:", "$env:username")
    
    Add-Content -Path $ScriptLogFile -Value 'Domain username not found'
	'Username being used is'+$username
}

Write-Host 'All actions performed by this script are written to the log file at ' $ScriptLogFile 

Add-Content -Path $ScriptLogFile -Value 'Username being used is: ' -NoNewline
Add-Content -Path $ScriptLogFile -Value  $($username.ToString())
Add-Content -Path $ScriptLogFile -Value 'Last run: ' -NoNewline
Add-Content -Path $ScriptLogFile -Value  $((Get-Date).ToString())

Add-Content -Path $ScriptLogFile -Value '========= TIME SYNC WITH SYDNEY DOMAIN ========='

#Syncs the time with our domain
try 
{
	net time /domain:sydney.ssw.com.au /set /y 
	Add-Content -Path $ScriptLogFile -Value 'Sydney time synced'
}
catch
{
	Add-Content -Path $ScriptLogFile -Value 'Sydney time sync failed'
}

Add-Content -Path $ScriptLogFile -Value '========= OFFICE TEMPLATES ========='

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Normal.dot'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Normal.dot'

try 
{
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination
    Add-Content -Path $ScriptLogFile -Value 'Normal.dot copied'
}
catch
{    
    Add-Content -Path $ScriptLogFile -Value 'Normal.dot copy failed'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Normal.dotm'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Normal.dotm'

try 
{
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'Normal.dotm copied'
}
catch
{    
    Add-Content -Path $ScriptLogFile -Value 'Normal.dotm copy failed'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/ProposalNormalTemplate.dotx'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\ProposalNormalTemplate.dotx'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'ProposalNormalTemplate.dotx copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'ProposalNormalTemplate.dotx copy failed'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/NormalEmail.dot'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\NormalEmail.dot'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'NormalEmail.dot copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'NormalEmail.dot copy failed'    
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Microsoft_Normal.dotx'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Microsoft_Normal.dotx'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'Microsoft_Normal.dotx copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'Microsoft_Normal.dotx copy failed'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Blank.potx'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Blank.potx'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'Blank.potx copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'Blank.potx copy failed'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/NormalEmail.dotm'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\NormalEmail.dotm'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'NormalEmail.dotm copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'NormalEmail.dotm copy failed'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/NormalEmail.dotm'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\QuickStyles\NormalEmail.dotm'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'NormalEmail.dotm copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'NormalEmail.dotm copy failed'
}

Add-Content -Path $ScriptLogFile -Value '========= OUTLOOK SIGNATURE ========='

$SignatureDestination  = $env:APPDATA + '\Microsoft\Signatures\'
New-Item -ItemType Directory -Force -Path $SignatureDestination  | Out-Null 

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Outlook/SSW_' + $username + '_Short_Default.htm'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Signatures\SSW.htm'

try {
    if (Test-Path $ScriptFileDestination) {
        Copy-Item $ScriptFileDestination -Destination ($ScriptFileDestination).Replace("SSW.htm","zzSSW.htm")
        Add-Content -Path $ScriptLogFile -Value 'Already found SSW.htm. Renaming it to zzSSW.htm'
    }
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'SSW.htm rename failed'
}
try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'SSW.htm copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'SSW.htm copy failed'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Outlook/SSW_' + $username + '_Short_Default.txt'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Signatures\SSW.txt'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'SSW.txt copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'SSW.txt copy failed'
}

$SignatureDestination  = $env:APPDATA + '\Microsoft\Signatures\SSW_files\'
New-Item -ItemType Directory -Force -Path $SignatureDestination  | Out-Null 

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Outlook/SSW_' + $username + '_Short_Default_files/colorschememapping.xml'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Signatures\SSW_files\colorschememapping.xml'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'colorschememapping.xml copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'colorschememapping.xml copy failed'

}
$ScriptFileSource = $ScriptTemplateSource + '/Templates/Outlook/SSW_' + $username + '_Short_Default_files/filelist.xml'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Signatures\SSW_files\filelist.xml'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'filelist.xml copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'filelist.xml copy failed'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Outlook/SSW_' + $username + '_Short_Default_files/themedata.thmx'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Signatures\SSW_files\themedata.thmx'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'themedata.thmx copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'themedata.thmx copy failed'
}

Add-Content -Path $ScriptLogFile -Value '========= SNAGIT THEME ========='

Stop-Process -name 'Snagit32', 'SnagitEditor', 'SnagitPriv'  -ErrorAction 'silentlycontinue'

$ScriptFileSource = $ScriptTemplateSource + '/Templates/SnagIt_DrawQuickStyles.xml'
$ScriptFileDestination = $env:APPDATA + '\..\Local\TechSmith\Snagit\DrawQuickStyles.xml'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value 'DrawQuickStyles.xml copied'
}
catch {
    Add-Content -Path $ScriptLogFile -Value 'DrawQuickStyles.xml copy failed'
}

#Opens up notepad at the end with our completed log
notepad C:\SSWTemplateScript_LastRun.log

