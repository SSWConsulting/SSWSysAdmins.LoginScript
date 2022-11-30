<#
.SYNOPSIS
    PowerShell Login Script for SSW.
.DESCRIPTION
    PowerShell Login Script for SSW.
    It flushes the DNS, copies Office templates from github to your machine.
.EXAMPLE
    PS> Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex (new-object net.webclient).downloadstring('https://github.com/SSWConsulting/SSWSysAdmins.LoginScript/raw/main/Script/SSWLoginScript.ps1')
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
2.6         Kaique Biancatti    28/04/2021      Added background popup, fixed urls to be "main" not master, fixed SysAdmin names
2.7         Kaique Biancatti    15/09/2021      Changed references from flea to sydfilesp03.sydney.ssw.com.au, changed SysAdmin names, disabled most signature fetches as we are moving to CodeTwo
2.8         Kaique Biancatti    02/11/2021      Deleted signature steps and error handling, we are fully using CodeTwo now so no need for signatures
2.9         Kaique Biancatti    27/01/2022      Deleted function to replace background, updated Intranet link
3.0         Kaique Biancatti    30/11/2022      Deleted admin check, deleted Sydney time sync, deleted domain account check, changed login folder location, refactored some log commands
3.1         Kaique Biancatti    30/11/2022      Deleted server log write (this scrips assumes you can be anywhere in the world, not connected to the domain), changed descriptions
3.2         Kaique Biancatti    30/11/2022      Added a stopwatch, deleted some junk from the folders

DO NOT FORGET TO UPDATE THE $ScriptVersion AND $ScriptLastUpdated VARIABLE BELOW
#>
param (
    [string]$username = ''
)
#Sets our Script version. Please update this variable anytime a new version is made available
$ScriptVersion = '3.1'

#Sets our last update date. Please update this variable anytime a new version is made available
$ScriptLastUpdated = "30/11/2022"

#Functions
#This function adds the error message to the log if any
Function Add-ErrorToLog {
    $RecentError = $Error[0]
    if ($RecentError -ne $null) {
        Add-Content -Path $ScriptLogFile -Value "   >> $($RecentError)"
    }
    else {
    }
}

#Let's time this!
$Script:Stopwatch = [system.diagnostics.stopwatch]::StartNew()

#Setting Github as the only place to get Templates from
Set-Variable -Name 'ScriptTemplateSource' -Value 'https://github.com/SSWConsulting/SSWSysAdmins.LoginScript/raw/main/'

#Initializing the LogFile
Set-Variable -Name 'ScriptLogFile' -Value "$Env:Temp\SSWLoginScript_LastRun.log"
Set-Content -Path $ScriptLogFile -Value 'SSWLoginScript Log' -Force
Add-Content -Path $ScriptLogFile -Value ''
Add-Content -Path $ScriptLogFile -Value 'Thanks. The Login Script is now finished!'

Write-Host 'This PowerShell script copies SSW Template Files from' $ScriptTemplateSource 'to your %AppData%\Microsoft\Templates\ folder'
Write-Host 'Please make sure that Word, Powerpoint and Outlook are closed. Open templates will not be replaced'

#This command is the same as ipconfig/flushdns, clears the DNS cache on the client
Clear-DnsClientCache

#This sets the security protocol to use all TLS versions. Without this, Powershell will use TLS1.0 which GitHub does not accept.
[Net.ServicePointManager]::SecurityProtocol = "tls12, tls11, tls"

Write-Host ''
Write-Host 'All actions performed by this script are written to the log file at' $ScriptLogFile

#Explains what this script does
Add-Content -Path $ScriptLogFile -Value ''
Add-Content -Path $ScriptLogFile -Value 'What did this script do?'
Add-Content -Path $ScriptLogFile -Value '   1. Flushed DNS'
Add-Content -Path $ScriptLogFile -Value '   2. Copied Office Templates from GitHub to your machine, as per the rule https://rules.ssw.com.au/have-a-companywide-word-template'
Add-Content -Path $ScriptLogFile -Value ''
Add-Content -Path $ScriptLogFile -Value '   Please review the success or failure below and errors if any:'

#Starts copying the office templates and signatures
$ScriptFileSource = $ScriptTemplateSource + '/Templates/Normal.dot'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Normal.dot'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination
    Add-Content -Path $ScriptLogFile -Value '   Normal.dot Copy                                 [Done]'
}
catch {    
    Add-Content -Path $ScriptLogFile -Value '   Normal.dot Copy(Word Open)                      [Failed]'
    Add-ErrorToLog
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Normal.dotm'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Normal.dotm'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   Normal.dotm Copy                                [Done]'
}
catch {    
    Add-Content -Path $ScriptLogFile -Value '   Normal.dotm Copy(Word Open)                     [Failed]'
    Add-ErrorToLog
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/ProposalNormalTemplate.dotx'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\ProposalNormalTemplate.dotx'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   ProposalNormalTemplate.dotx Copy                [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   ProposalNormalTemplate.dotx Copy(Word Open)     [Failed]'
    Add-ErrorToLog
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/NormalEmail.dot'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\NormalEmail.dot'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dot Copy                            [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dot Copy(Word Open)                 [Failed]'
    Add-ErrorToLog
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Microsoft_Normal.dotx'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Microsoft_Normal.dotx'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   Microsoft_Normal.dotx Copy                      [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   Microsoft_Normal.dotx Copy(Word Open)           [Failed]'
    Add-ErrorToLog
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/Blank.potx'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\Blank.potx'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   Blank.potx Copy                                 [Done]'
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   Blank.potx Copy(Word Open)                      [Failed]'
    Add-ErrorToLog
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/NormalEmail.dotm'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\Templates\NormalEmail.dotm'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Template Copy                  [Done]'
}
catch [System.IO.DirectoryNotFoundException] {
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Template Copy(Path Not Found)  [Failed]'
    Add-ErrorToLog
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Template Copy(Outlook Open)    [Failed]'
    Add-ErrorToLog
}

$ScriptFileSource = $ScriptTemplateSource + '/Templates/NormalEmail.dotm'
$ScriptFileDestination = $env:APPDATA + '\Microsoft\QuickStyles\NormalEmail.dotm'

try {
    Invoke-WebRequest -Uri $ScriptFileSource -OutFile $ScriptFileDestination 
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Quickstyle Copy                [Done]'
}
catch [System.IO.DirectoryNotFoundException] {
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Quickstyle Copy(Path Not Found)[Failed]'
    Add-ErrorToLog
}
catch {
    Add-Content -Path $ScriptLogFile -Value '   NormalEmail.dotm Quickstyle Copy(Outlook Open)  [Failed]'
    Add-ErrorToLog
}

Add-Content -Path $ScriptLogFile -Value ''
Add-Content -Path $ScriptLogFile -Value ''

#Shows the Script last update date in the Log
Add-Content -Path $ScriptLogFile -Value "   Version: $ScriptVersion (updated on $ScriptLastUpdated)"

#Shows the last time the script was run on in the Log
Add-Content -Path $ScriptLogFile -Value "   Last run on your computer: $((Get-Date).ToString())"
Add-Content -Path $ScriptLogFile -Value "   This script took $($Script:Stopwatch.Elapsed.ToString('mm')) minutes and $($Script:Stopwatch.Elapsed.ToString('ss')) seconds to run"
Add-Content -Path $ScriptLogFile -Value ''
Add-Content -Path $ScriptLogFile -Value 'From your friendly SysAdmins'
Add-Content -Path $ScriptLogFile -Value 'Kiki Biancatti & Warwick Leahy & Chris Schultz & Mehmet Ozdemir & Ash Anil'
Add-Content -Path $ScriptLogFile -Value 'https://sswcom.sharepoint.com/sites/SSWSysAdmins'

#Let's stop timing this!
$Script:Stopwatch.Stop();

#Opens up notepad at the end with our completed log
notepad $ScriptLogFile