<#
.SYNOPSIS
    Monthly Public Folders

.DESCRIPTION
    Creates public folders per month of a year and fitting transport rules to sort them correctly.

.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does

.INPUTS
    None

.OUTPUTS
    None

.NOTES
    Author: Michael Schönburg
    GitHub Repository: https://github.com/MichaelSchoenburg/MonthlyPublicFolders

    This projects code loosely follows the PowerShell Practice and Style guide*, 
    as well as Microsofts PowerShell scripting performance considerations**.

    * https://poshcode.gitbook.io/powershell-practice-and-style/

    ** https://docs.microsoft.com/en-us/powershell/scripting/dev-cross-plat/performance/script-authoring-considerations?view=powershell-7.1
#>

[CmdletBinding()]
param (
    # General
    [Parameter(Mandatory)]
    [ValidateSet(
        "DE",
        "EN"
    )]
    [string]
    $Language,

    # Who is the parent?
    [Parameter(Mandatory)]
    [string]
    $ParentPublicFolderPath,
    
    [Parameter(Mandatory)]
    [string]
    [ValidateScript({$_ -like "*@*.*"})]
    $ParentPublicFolderMailAddress,
    
    # For new mail addresses
    [Parameter(Mandatory)]
    [ValidateScript({$_ -like "*.*"})]
    [string]
    $Domain,

    # Auth
    [string]
    $AppID = 'd1186226-581c-44e6-a96b-78d7b90cc8cf',

    [string]
    $CertName = 'OnboardingScript',

    [string]
    $PathToCert = 'Cert:\LocalMachine\My',

    [string]
    $DelegatedOrg
)

#region FUNCTIONS

function Write-ConsoleLog {
    <#
    .SYNOPSIS
    Logs an event to the console.
    
    .DESCRIPTION
    Writes text to the console with the current date (US format) in front of it.
    
    .PARAMETER Text
    Event/text to be outputted to the console.
    
    .EXAMPLE
    Write-ConsoleLog -Text 'Subscript XYZ called.'
    
    Long form

    .EXAMPLE
    Log 'Subscript XYZ called.
    
    Short form
    #>
    [alias('Log')]
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
        Position = 0)]
        [string]
        $Text
    )

    # Save current VerbosePreference
    $VerbosePreferenceBefore = $VerbosePreference

    # Enable verbose output
    $VerbosePreference = 'Continue'

    # Write verbose output
    Write-Verbose "$( Get-Date -Format 'MM/dd/yyyy HH:mm:ss' ) - $( $Text )"

    # Restore current VerbosePreference
    $VerbosePreference = $VerbosePreferenceBefore
}

function New-MonthlyPublicFolder {
    param (
        [ValidateScript({
            if ($_ -ne "") {
                $true
            } else {
                throw "$_ is empty."
            }
        })]
        [string]
        $Name,

        [ValidateScript({
            if ($_ -ne "") {
                $true
            } else {
                throw "$_ is empty."
            }
        })]
        [string]
        $MailAddress,

        [ValidateScript({
            if ($_ -ne "") {
                $true
            } else {
                throw "$_ is empty."
            }
        })]
        [string]
        $ParentPublicFolder,

        [ValidateScript({
            if ($_ -ne "") {
                $true
            } else {
                throw "$_ is empty."
            }
        })]
        [string]
        $FullPath
    )
    
    $pf = New-PublicFolder -Name $Name -Path $ParentPublicFolder -Confirm:$false

    # Set up mail address for the public folder
    Enable-MailPublicFolder -Identity $pf.Identity
    Set-MailPublicFolder -Identity $pf.Identity -EmailAddressPolicyEnabled:$false -EmailAddresses $MailAddress

    # Add permissions to be able to receive mails
    Remove-PublicFolderClientPermission $pf.Identity -User Anonym -Confirm:$false
    Add-PublicFolderClientPermission $pf.Identity -User Anonym -AccessRights CreateItems
}

#endregion FUNCTIONS
#region DECLARATIONS

# Dates
$Year = (Get-Date).Year
$CurrentMonth = (Get-Date).Month
if ($CurrentMonth.ToString().Length -eq 1){$CurrentMonth = "0$( $CurrentMonth )"}

# Basic path correction
if ($ParentPublicFolderPath -notlike "\*") {$ParentPublicFolderPath = "\$( $ParentPublicFolderPath )"}

# Auth
$Thumbprint = ((Get-ChildItem $PathToCert).Where({$_.Subject -eq 'CN=onboardingScript'})).Thumbprint

#endregion DECLARATIONS
#region EXECUTION

if ($DelegatedOrg) {
    Connect-ExchangeOnline -DelegatedOrganization $DelegatedOrg -AppId $AppID -CertificateThumbprint $Thumbprint -ShowBanner:$false
} else {
    Connect-ExchangeOnline -AppId $AppID -CertificateThumbprint $Thumbprint -ShowBanner:$false
}

# Clean up any old rules
$OldRules = (Get-TransportRule).where({($_.Name -like "MPF # *") -and ($_.Name -notlike "MPF # $( $ParentPublicFolderMailAddress ) `# $( $Year ) `#*")})
$OldRules | Remove-TransportRule

# Check if already created
if ((Get-TransportRule).where({$_.Name -like "MPF # $( $ParentPublicFolderMailAddress ) # $( $Year ) #  $( $CurrentMonth ) *"})) {
    $PublicFolderForYear = New-PublicFolder -Name "$( $Year )" -Path $ParentPublicFolderPath
    $ParentPublicFolderPath = $PublicFolderForYear.FullPath

    $Numbers = 1..12
    
    ForEach ($Number in $Numbers)
    {
        $Months = Get-Content ".\Months_$( $Language ).txt"
        $Month = $Months[$Number]
        
        # Prepend 0 if single digit number
        if ($Number.Length -eq 1){$Number = "0$( $Number )"}
    
        #####################################
        # Folder for Regular Mails
        #####################################
        
        $HashArguments = @{
            ParentPublicFolder = $ParentPublicFolder
            FullPath    = "$( $ParentPublicFolder )\$( $PublicFolderName )"
            MailAddress = "$( $Number ).$( $Year )@$( $Domain )"
            Name = "$( $Number ) - $( $Month ) $( $Year )"
        }

        $RegularMailsFolder = New-MonthlyPublicFolder @HashArguments

        #####################################
        # Folder for System Notifications
        #####################################
    
        switch ($Language) {
            "DE" {
                $SystemNotificationFolderName = "Systembenachrichtigungen"
            }
            "EN" {
                $SystemNotificationFolderName = "SystemNotifications"
            }
        }

        $HashArguments = @{
            ParentPublicFolder = $RegularMailsFolder.FullPath
            FullPath    = "$( $ParentPublicFolder )\$( $PublicFolderName )"
            MailAddress = "$( $Number ).$( $Year ).$( $SystemNotificationFolderName )@$( $Domain )"
            Name = $SystemNotificationFolderName
        }

        $SystemNotificationFolder = New-MonthlyPublicFolder @HashArguments
    
        #####################################
        # Rule for Regular Mails
        #####################################
        
        $HashArguments = @{
            Name = "MPF # $( $ParentPublicFolderMailAddress ) # $( $Year ) # $( $Month )"
            SentTo = $ParentPublicFolderMailAddress
            RedirectMessageTo = "$( $Number ).$( $Year )@$( $Domain )"
            StopRuleProcessing = $true
            ActivationDate = (GetDate "01.$( $Number ).$( $Year ) 01:00:00")
            ExpiryDate = (GetDate "01.$([int]$Number + 1).$( $Year ) 00:00:00")
        }

        New-TransportRule @HashArguments

        #####################################
        # Rule for System Notifications
        #####################################

        $HashArguments = @{
            Name = "MPF # $( $ParentPublicFolderMailAddress ) # $( $Year ) # $( $Month ) # Sys"
            SentTo = $ParentPublicFolderMailAddress
            From = $ParentPublicFolderMailAddress
            RedirectMessageTo = "$( $Number ).$( $Year ).$( $SystemNotificationFolderName )@$( $Domain )"
            StopRuleProcessing = $true
            ActivationDate = (Get-Date "01.$( $Number ).$( $Year ) 01:00:00")
            ExpiryDate = (GetDate "01.$([int]$Number + 1).$( $Year ) 00:00:00")
        }

        New-TransportRule @HashArguments
    }    
}

# Disconnect Exchange Online
Get-PSSession | Remove-PSSession

#endregion EXECUTION
