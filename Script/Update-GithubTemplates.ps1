<#
.SYNOPSIS
    PowerShell GitHub template uploader.
.DESCRIPTION
    PowerShell GitHub template uploader.
    It downloads all Outlook templates from file server, commits them to a local git repo and uploads them to GitHub at https://github.com/SSWConsulting/SSWSysAdmins.LoginScript
.EXAMPLE
    This script is triggered on a schedule using the Windows Task Scheduler.
.INPUTS
    Configuration file: Config.psd1
.OUTPUTS
    Uploads new Outlook signatures to GitHub.
.NOTES
    It will overwrite all Outlook certificates on GitHub.

    Created by Kaique "Kiki" Biancatti for SSW.
#>

# Start a stopwatch
$Script:Stopwatch = [system.diagnostics.stopwatch]::StartNew()

# Importing the configuration file
$config = Import-PowerShellDataFile $PSScriptRoot\Config.PSD1

# Creating variables to determine magic strings and getting them from the configuration file
$LogFile = $config.LogFile
$OriginEmail = $config.OriginEmail
$TargetEmail = $config.TargetEmail
$LogModuleLocation = $config.LogModuleLocation
$SourceCopyFolder = $config.SourceCopyFolder
$Script:ErrorFlag = $false

# Importing the SSW Write-Log module
Import-Module -Name $LogModuleLocation

<#
.SYNOPSIS
Function to download new Outlook signatures from on-premises file server and auto-upload them to GitHub on a schedule (using Task Scheduler).

.DESCRIPTION
Function to download new Outlook signatures from on-premises file server and auto-upload them to GitHub on a schedule (using Task Scheduler).
Uses robocopy to filter the necessary files only and git commands to auto-upload to Git.

.PARAMETER LogFile
The location of the logfile.

.PARAMETER LogModuleLocation
The location of the module to write logs.

.PARAMETER SourceCopyFolder
The source folder with the Outlook templates.

.EXAMPLE
PS> Update-GithubTemplates -LogFile $LogFile -LogModuleLocation $LogModuleLocation -SourceCopyFolder $SourceCopyFolder
#>
function Update-GithubTemplates {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $LogFile,
        [Parameter(Mandatory)]
        $LogModuleLocation,
        [Parameter(Mandatory)]
        $SourceCopyFolder
    )

    try {
        $ParentFolder = (Split-Path -Path $PSScriptRoot -Parent) + "\Templates"
        robocopy $SourceCopyFolder $ParentFolder Blank.potx Microsoft_Normal.dotx Normal.dot Normal.dotm Normal.dotx NormalEmail.dot NormalEmail.dotm NormalEmail.dotx ProposalNormalTemplate.dotx SSW.snagtheme White-SSW-Wallpaper.bmp /copyall /is
        Write-Log -File $LogFile -Message "Successfully copied items from $SourceCopyFolder to $ParentFolder..."
    }
    catch {
        $RecentError = $Error[0]
        $Script:ErrorFlag = $true
        Write-Log -File $LogFile -Message "ERROR on function Update-GithubTemplates, could not copy items from $SourceCopyFolder - $RecentError"
    }
    try {
        $SourceCopyFolderOutlook = $SourceCopyFolder + "\Outlook"
        $ParentFolderOutlook = $ParentFolder + "\Outlook"
        robocopy $SourceCopyFolderOutlook $ParentFolderOutlook SSW_* colorschememapping.xml filelist.xml themedata.thmx /E /xd temp
        Write-Log -File $LogFile -Message "Successfully copied items from $SourceCopyFolderOutlook to $ParentFolderOutlook..."
    }
    catch {
        $RecentError = $Error[0]
        $Script:ErrorFlag = $true
        Write-Log -File $LogFile -Message "ERROR on function Update-GithubTemplates, could not copy items from $SourceCopyFolderOutlook - $RecentError"
    }
    try {
        #git status
        git add ../Templates/Outlook/
        Write-Log -File $LogFile -Message "Sucessfully added /Templates/Outlook/ to Git..."
    }
    catch {
        $RecentError = $Error[0]
        $Script:ErrorFlag = $true
        Write-Log -File $LogFile -Message "ERROR on function Update-GithubTemplates, could not add /Templates/Outlook/ to Git - $RecentError"
    }
    try {
        $Today = get-date -f dd/MM/yyy 
        $GitMessage = git commit -m "Auto commit on $Today to have GitHub up-to-date with on-premises fileserver" 
        Write-Log -File $LogFile -Message "Sucessfully commited to git - $GitMessage"
    }
    catch {
        $RecentError = $Error[0]
        $Script:ErrorFlag = $true
        Write-Log -File $LogFile -Message "ERROR on function Update-GithubTemplates, could not auto commit to Git - $RecentError"
    }
    try {
        $GitMessage = git push origin master
        Write-Log -File $LogFile -Message "Successfully push new files to Git..."
    }
    catch {
        $RecentError = $Error[0]
        $Script:ErrorFlag = $true
        Write-Log -File $LogFile -Message "ERROR on function Update-GithubTemplates, could not auto push to Git - $RecentError"
    }
    
}

<#
.SYNOPSIS
Function to build an email and send it in case any errors pop up.

.DESCRIPTION
Function to build an email and send it in case any errors pop up.
Only send the email if errors pop up.

.PARAMETER LogFile
The location of the logfile.

.PARAMETER TargetEmail
The email that the message will be sent to.

.PARAMETER OriginEmail
The email that the message will originate from.

.EXAMPLE
PS> Send-Email -LogFile $LogFile -TargetEmail $TargetEmail -OriginEmail $OriginEmail
#>
function Send-Email {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0, Mandatory = $true)]
        [string] $LogFile,
        [Parameter(Position = 1, Mandatory = $true)]
        [string] $TargetEmail,
        [Parameter(Position = 2, Mandatory = $true)]
        [string] $OriginEmail
 
    )
 
    try {
        # Let's create the HTML body of the email
        $Body = @"
         <div style='font-family:Calibri;'>
         <p>There was an error in the Update-GitHubTemplates script.</p>
         Check the log file to see what the error was (look for ERROR) in <a href=$LogFile> $LogFile</a>
         You could also run the script manually in an Admin PowerShell, script is in \\$env:computername\$PSScriptRoot<br>
         This script took $($Script:Stopwatch.Elapsed.Seconds) seconds to run.</p>
         <p>-- Powered by SSW.LoginScript<br></p>
         <p>
         GitHub: <a href=https://github.com/SSWConsulting/SSWSysAdmins.LoginScript>SSWSysAdmins.LoginScript</a><br>
         Server: $env:computername <br>
         Folder: $PSScriptRoot</p>
"@
 
        if ($Script:ErrorFlag -eq $true) { 
            Send-MailMessage -from $OriginEmail -to $TargetEmail -Subject "SSW.LoginScript - Error on Auto Upload of Templates - Further manual action required" -Body $body -SmtpServer "ssw-com-au.mail.protection.outlook.com" -bodyashtml
            Write-Log -File $LogFile -Message "Succesfully sent email to $TargetEmail from $OriginEmail..."
        }
        else {
            Write-Log -File $LogFile -Message "Succesfully skipped send email, no errors..."
        }        
    }
    catch {
        $RecentError = $Error[0]
        Write-Log -File $LogFile -Message "ERROR sending email to $TargetEmail from $OriginEmail - $RecentError"
    }  
}

Update-GithubTemplates -LogFile $LogFile -LogModuleLocation $LogModuleLocation -SourceCopyFolder $SourceCopyFolder
Send-Email -LogFile $LogFile -TargetEmail $TargetEmail -OriginEmail $OriginEmail

# Let's stop timing this!
$Script:Stopwatch.Stop();