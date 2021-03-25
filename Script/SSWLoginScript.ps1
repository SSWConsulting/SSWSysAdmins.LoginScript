<#
.SYNOPSIS
    PowerShell Login Script for SSW.
.DESCRIPTION
    PowerShell Login Script for SSW.
    It checks if running elevated, flushes the DNS, syncs PC time with Sydney server, copies Office templates and outlook signatures from fileserver or github to your machine, and copies snagit templates.
.EXAMPLE
    PS> iex (new-object net.webclient).downloadstring('https://github.com/SSWConsulting/SSWSysAdmins.LoginScript/raw/master/Script/SSWLoginScript.ps1')
.OUTPUTS
    flushes the DNS, syncs PC time with Sydney server, copies Office templates and outlook signatures from fileserver or github to your machine, and copies snagit templates.
.NOTES
Version     Author              Date            Comment
1.0         Greg Harris         12/03/2018      Initial Version - Based on SSWLoginScript.bat
1.1         Kaique Biancatti    07/06/2018      Added the correct link to GitHub and added TLS options to connect to HTTPS. Also added name prompt.
1.2         Kaique Biancatti    08/06/2018      Added self elevation of PowerShell script, comments, backup logic, and reorganizing of code.
1.3	        Kaique Biancatti    16/07/2018      Added open notepad with log at the end of script.
1.4         Kaique Biancatti    17/07/2018      Added time sync with Sydney server.
1.5         Greg Harris         17/07/2018      Added FlushDNS, check for file share and use if available. If not, use github.
1.6         Kaique Biancatti    20/07/2018      Added LogWrite function to write logs in our fileserver for debugging.
1.7         Kaique Biancatti    20/07/2018      Changed some log entries. Rearranged the code to look better.
1.8         Kaique Biancatti    25/07/2018      Changed InputBox text. Added ScriptVersion variable. Changed how the log looks.
1.9         Kaique Biancatti    31/07/2018      Changed some log messages. Fixed some typos.
2.0         Kaique Biancatti    27/08/2018      Changed all TemplateScript names to LoginScript. 
2.1         Kaique Biancatti    13/09/2018      Changed InputBox description. Changed LogFile structure.
2.2         Kaique Biancatti    12/06/2019      Changed Fail Messages in Log explaining what might be the reason. Changed script name to SSWLoginScript. Added ".sydney.ssw.com.au" to fileserver path.
2.3         Alex Breskin        28/08/2019      Added logging conditions for directories that may not exist
2.4         Kaique Biancatti    24/09/2019      Added SSW background for domain-joined computers
2.5         Kaique Biancatti    17/07/2020      Changed log function name, changed the GitHub URL, cleaned up the code a bit, added Comment-Based help at the top

DO NOT FORGET TO UPDATE THE SCRIPTVERSION VARIABLE BELOW
#>

param (
    [string]$username = ''
)

#Sets our Script version. Please update this variable anytime a new version is made available
$ScriptVersion = '2.5'

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
If ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) -eq $False) {
    # Relaunch as an elevated process:
    Start-Process powershell.exe "-File", ('"{0}"' -f $MyInvocation.MyCommand.Path) -Verb RunAs
    exit
}
#This bit will create a function to write a log in our fileserver
$Logfile = "\\flea\DataSSW\DataSSWEmployees\LoginScriptUserLogs.log"

Function Write-Log {
    $username = $env:USERNAME   
    $PcName = $env:computername
    $Stamp = (Get-Date).toString("dd/MM/yyyy HH:mm:ss")
    $Line = "$Stamp $PcName $Username"
    Add-content $Logfile -value $Line
}

#This part tests if we can find the fileserver. If we can't, we will get the Signatures and Templates from GitHub instead.
$ShareExists = Test-Path $('filesystem::\\fileserver\DataSSW\DataSSWEmployees\Templates')
if ($ShareExists -eq $true) {
    Set-Variable -Name 'ScriptTemplateSource' -Value 'file://fileserver.sydney.ssw.com.au/DataSSW/DataSSWEmployees'
}
else {
    Set-Variable -Name 'ScriptTemplateSource' -Value 'https://github.com/SSWConsulting/SSWSysAdmins.LoginScript/raw/main/'
}

#Initializing the LogFile
Set-Variable -Name 'ScriptLogFile' -Value 'C:\SSWLoginScript_LastRun.log'
Set-Content -Path $ScriptLogFile -Value 'SSWLoginScript Log' -Force
Add-Content -Path $ScriptLogFile -Value ''
Add-Content -Path $ScriptLogFile -Value 'Thanks. The Login Script is now finished!'

Write-Host 'This PowerShell script copies SSW Template Files from' $ScriptTemplateSource 'to your %AppData%\Microsoft\Templates\ folder'
Write-Host 'Please make sure that Word, Powerpoint and Outlook are closed. Open templates will not be replaced'

#This command is the same as ipconfig/flushdns, clears the DNS cache on the client
Clear-DnsClientCache

#This sets the security protocol to use all TLS versions. Without this, Powershell will use TLS1.0 which GitHub does not accept.
[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

#This bit tries to find if the user is running on a domain account.
$domain = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[0]
if ($domain -eq 'SSW2000') {
    if ($username -eq '') {
        $username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[1]
    }
    $noDomainUsername = $false
}

if ($username -eq '') {
    Write-Host ''
    Write-Host 'Username parameter required if not run on a SSW2000 domain account. Please input username on the pop-up box that just appeared somewhere in your screen.'

    #Calling a VB Prompt for the user if there is no username set    
    [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    $username = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your name as FirstLast`nWARNING: Case Sensitive eg. AdamCogan", "Please input your SSW username:", "AdamCogan")
    $noDomainUsername = $true
}

Write-Host ''
Write-Host 'All actions performed by this script are written to the log file at' $ScriptLogFile
Write-Host 'You can also find who ran this script in' $LogFile

#Explains what this script does
Add-Content -Path $ScriptLogFile -Value ''
Add-Content -Path $ScriptLogFile -Value 'What did this script do?'
Add-Content -Path $ScriptLogFile -Value '   1. Flushed DNS'
Add-Content -Path $ScriptLogFile -Value '   2. Synchronized your PC time with the computer time of the Sydney server'
Add-Content -Path $ScriptLogFile -Value '   3. Copied Office Templates to your machine, as per the rule https://rules.ssw.com.au/have-a-companywide-word-template'
Add-Content -Path $ScriptLogFile -Value '     a. If you do not have access to our fileserver, copied them from GitHub'
Add-Content -Path $ScriptLogFile -Value '   4. Copied Outlook Signatures to your PC (using the same rules as above)'
Add-Content -Path $ScriptLogFile -Value '   5. Closed SnagIt if it was open, and copied its templates to your PC (using the same rules as above)'
Add-Content -Path $ScriptLogFile -Value ''
Add-Content -Path $ScriptLogFile -Value '   Please review the success or failure below:'

#Syncs the time with our domain
try {
    net time /domain:sydney.ssw.com.au /set /y 
    Add-Content -Path $ScriptLogFile -Value '   Sydney Time Sync                           [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   Sydney Time Sync(No access to server)      [Failed]'
}

#Starts copying the office templates and signatures
$ScriptFileSource = $ScriptTemplateSource + '/Templates/Normal.dot'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Normal.dot'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination
    Add-Content -Path $ScriptLogFile -Value '   Normal.dot Copy                            [Done]'
}
catch {    
    Add-Content -Path $ScriptLogFile -Value '   Normal.dot Copy(Word Open)                 [Failed]'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Normal.dotm'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Normal.dotm'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   Normal.dotm Copy                           [Done]'
}
catch {    
    Add-Content -Path $ScriptLogFile -Value '   Normal.dotm Copy(Word Open)                [Failed]'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/ProposalNormalTemplate.dotx'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\ProposalNormalTemplate.dotx'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   ProposalNormalTemplate.dotx Copy           [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   ProposalNormalTemplate.dotx Copy(Word Open)[Failed]'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/NormalEmail.dot'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\NormalEmail.dot'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dot Copy                       [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dot Copy(Word Open)            [Failed]'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Microsoft_Normal.dotx'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Microsoft_Normal.dotx'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   Microsoft_Normal.dotx Copy                 [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   Microsoft_Normal.dotx Copy(Word Open)      [Failed]'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Blank.potx'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Blank.potx'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   Blank.potx Copy                            [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   Blank.potx Copy(Word Open)                 [Failed]'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/NormalEmail.dotm'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\NormalEmail.dotm'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Copy                      [Done]'
}
catch [System.IO.DirectoryNotFoundException] {
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Copy(Path Not Found)      [Failed]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Copy(Outlook Open)        [Failed]'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/NormalEmail.dotm'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\QuickStyles\NormalEmail.dotm'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Copy                      [Done]'
}
catch [System.IO.DirectoryNotFoundException] {
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Copy(Path Not Found)      [Failed]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Copy(Outlook Open)        [Failed]'
}

$SignatureDestination = $env:APPDATA + '\Microsoft\Signatures\'
New-Item -ItemType Directory -Force -Path $SignatureDestination  | Out-Null 

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Outlook/SSW_' + $username + '_Short_Default.htm'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Signatures\SSW.htm'

try {
    if (Test-Path $ScriptFileDestination) {
        Copy-Item $ScriptFileDestination -Destination ($ScriptFileDestination).Replace("SSW.htm", "zzSSW.htm")
        Add-Content -Path $ScriptLogFile -Value '   SSW.htm Signature Copy or Replace          [Replaced]'
    }
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   SSW.htm Signature Copy(Outlook Open)       [Failed]'
}
try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   SSW.htm Signature Copy                     [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   SSW.htm Signature Copy(Outlook Open)       [Failed]'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Outlook/SSW_' + $username + '_Short_Default.txt'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Signatures\SSW.txt'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   SSW.txt Signature Copy                     [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   SSW.txt Signature Copy(Outlook Open)       [Failed]'
}

$SignatureDestination = $env:APPDATA + '\Microsoft\Signatures\SSW_files\'
New-Item -ItemType Directory -Force -Path $SignatureDestination  | Out-Null 

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Outlook/SSW_' + $username + '_Short_Default_files/colorschememapping.xml'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Signatures\SSW_files\colorschememapping.xml'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   colorschememapping.xml Signature Copy      [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   colorschememapping.xml Signature Copy      [Failed]'

}
$ScriptFileSource = $ScriptTemplateSource + '/Templates/Outlook/SSW_' + $username + '_Short_Default_files/filelist.xml'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Signatures\SSW_files\filelist.xml'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   filelist.xml Signature Copy                [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   filelist.xml Signature Copy(No User)       [Failed]'
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Outlook/SSW_' + $username + '_Short_Default_files/themedata.thmx'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Signatures\SSW_files\themedata.thmx'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   themedata.thmx Signature Copy              [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   themedata.thmx Signature Copy(No User)     [Failed]'
}

#We need admin permissions to do this. If log stops here, it is because we have no privileges
Stop-Process -name 'Snagit32', 'SnagitEditor', 'SnagitPriv'  -ErrorAction 'silentlycontinue'

$ScriptFileSource = $ScriptTemplateSource + '/Templates/SnagIt_DrawQuickStyles.xml'
$ScriptFileDestination = $env:APPDATA + '\..\Local\TechSmith\Snagit\DrawQuickStyles.xml'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   DrawQuickStyles.xml Copy                   [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   DrawQuickStyles.xml Copy(SnagIt Open)      [Failed]'
}

#Set the computer background as SSW's image (for Domain-Joined only)
function Set-WallPaper([string]$desktopImage) {
    set-itemproperty -path "HKCU:Control Panel\Desktop" -name WallPaper -value $desktopImage
    RUNDLL32.EXE USER32.DLL, UpdatePerUserSystemParameters , 1 , True
}

#Download the SSW wallpaper from GitHub
$mydocuments = [environment]::getfolderpath("mydocuments")
$mydocumentsfull = $mydocuments + "\SSWBackground.bmp"
$url = "https://github.com/SSWConsulting/SSWSysAdmins.LoginScript/raw/master/Script/White-SSW-Wallpaper.bmp"
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($url, $mydocumentsfull)
Set-Wallpaper $mydocumentsfull

#If computer is domain-joined, set the wallpaper
if ($noDomainUsername -eq $False) {
    Set-Wallpaper $mydocumentsfull
    Add-Content -Path $ScriptLogFile -Value '   SSWBackground.jpg Copy                     [Done]'
}

#Writes the log in our server
Write-Log

Add-Content -Path $ScriptLogFile -Value ''
Add-Content -Path $ScriptLogFile -Value ''

#Shows the Script Version in the Log
Add-Content -Path $ScriptLogFile -Value '   Version: ' -NoNewline
Add-Content -Path $ScriptLogFile -Value $ScriptVersion

#Shows the last time the script was run on in the Log
Add-Content -Path $ScriptLogFile -Value '   Last run: ' -NoNewline
Add-Content -Path $ScriptLogFile -Value  $((Get-Date).ToString())

if ($noDomainUsername -eq $false) {

    Add-Content -Path $ScriptLogFile -Value '   Domain username: ' -NoNewline
    Add-Content -Path $ScriptLogFile -Value  $($username.ToString())
}	
else {
    Add-Content -Path $ScriptLogFile -Value '   Domain username: not found'
    Add-Content -Path $ScriptLogFile -Value '   Manual username: ' -NoNewline
    Add-Content -Path $ScriptLogFile -Value  $($username.ToString())
}

Add-Content -Path $ScriptLogFile -Value ' '
Add-Content -Path $ScriptLogFile -Value 'From your friendly System Administrators'
Add-Content -Path $ScriptLogFile -Value 'Steven Andrews & Kiki Biancatti & Mehmet Ozdemir'
Add-Content -Path $ScriptLogFile -Value 'sswcom.sharepoint.com/SysAdmin'

#Opens up notepad at the end with our completed log
notepad C:\SSWLoginScript_LastRun.log

