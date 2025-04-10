Function New-PSModuleScaffolding
{
    <#
    .SYNOPSIS
    Generates new PowerShell module files.

    .DESCRIPTION
    Generates new PowerShell module files.

    .PARAMETER ModuleName
    The name of the module to be created. File names and module manifest contents will match the value provided.

    .PARAMETER Description
    A brief description of the functionality provided by the module. Module manifest contents will match the value provided.

    .PARAMETER Author
    The name of the primary author.

    .PARAMETER CompanyName
    The name of the company with which the author is employed.

    .PARAMETER Copyright
    Copyright statement.

    .PARAMETER Guid
    A globally unique identifier.

    .PARAMETER ModuleVersion
    The version of code the module is to be designated as by the author.

    .PARAMETER Path
    The folder path you wish to create the module in. Do not include file names; .psm1 and .psd1 files are automatically named based on the ModuleName parameter value.

    .EXAMPLE
    New-Module -ModuleName FooBar -Description "A nonsense module."

    .EXAMPLE
    New-Module -ModuleName FooBar -Description "A nonsense module." -Guid $([guid]::NewGuid().Guid)

    .EXAMPLE
    New-Module -ModuleName FooBar -Description "A nonsense module. New and improved!" -ModuleVersion 2.0

    .EXAMPLE
    New-Module -ModuleName FooBar -Description "A nonsense module." -Author "David Whittaker" -Company "PrivateCorpLLC"
    #>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Position=0, Mandatory=$True, ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullOrEmpty()]
        [string]$ModuleName,

        [Parameter(Position=1, Mandatory=$True, ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullOrEmpty()]
        [string]$Description,

        [Parameter(Position=1, Mandatory=$True, ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullOrEmpty()]
        [string]$Author,

        [Parameter(Mandatory=$False, ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullOrEmpty()]
        [string]$CompanyName = $Author,

        [Parameter(Mandatory=$False, ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullOrEmpty()]
        [string]$Copyright = "",

        [Parameter(Mandatory=$False, ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullOrEmpty()]
        [string]$Guid = [guid]::NewGuid().Guid,

        [Parameter(Mandatory=$False, ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullOrEmpty()]
        [string]$ModuleVersion = '1.0.0',

        [Parameter(Mandatory=$False, ValueFromPipelineByPropertyName=$True)]
        [ValidateNotNullOrEmpty()]
        [string]$Path = $PWD
    )

    BEGIN
    {
        #region Staging
        # Create Variables
        $Result = @()

        $ManifestParameters = @{
            RootModule        = ".\$ModuleName.psm1"
            Description       = $Description
            Author            = $Author
            CompanyName       = $CompanyName
            Copyright         = $Copyright
            Guid              = $GUID
            ModuleVersion     = $ModuleVersion
            Path              = "$Path\$ModuleName\$ModuleName.psd1"
        }
        #endregion Staging
    }

    PROCESS
    {
        #region Main Workflow
        # Make directories and files
        $Null = New-Item $Path\$ModuleName\Public -ItemType Directory
        $Null = New-Item $Path\$ModuleName\Private -ItemType Directory
        $Null = New-Item $Path\$ModuleName\$ModuleName.psm1

        Get-Content -Path "$PSScriptRoot/../Templates/Template_Module.ps1" -Raw | Out-File $Path\$ModuleName\$ModuleName.psm1 -Encoding ascii

        # Generate module manifest
        New-ModuleManifest @ManifestParameters

        $Result = Get-Item -Path $Path\$ModuleName

        #endregion Main Workflow
    }

    END
    {
        # Output results at the end of the command
        Return $Result
    }
}