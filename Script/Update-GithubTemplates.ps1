
# Importing the configuration file
$config = Import-PowerShellDataFile $PSScriptRoot\Config.PSD1

# Creating variables to determine magic strings and getting them from the configuration file
$LogFile = $config.LogFile
$OriginEmail = $config.OriginEmail
$TargetEmail = $config.TargetEmail
$LogModuleLocation = $config.LogModuleLocation
$SourceCopyFolder = $config.SourceCopyFolder

# Importing the SSW Write-Log module
Import-Module -Name $LogModuleLocation

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
        Write-Log -File $LogFile -Message "ERROR on function Update-GithubTemplates, could not copy items from $SourceCopyFolderOutlook - $RecentError"
    }
    
}

Update-GithubTemplates -LogFile $LogFile -LogModuleLocation $LogModuleLocation -SourceCopyFolder $SourceCopyFolder