Function New-PSScriptScaffolding
{
    <#
        .SYNOPSIS
        Creates a new PowerShell script.

        .DESCRIPTION
        Creates a new PowerShell script and outputs the results to either the console or a file.

        .PARAMETER Author
        The name of the script's author.

        .PARAMETER CompanyName
        The company name of the script's author.

        .PARAMETER OutputPath
        The path to which to export the script. Absolute and relative paths are both supported.

        .PARAMETER Force
        Forces an overwrite of an existing file.

        .PARAMETER PassThru
        Returns output to the console in addition to producing a file.

        .EXAMPLE
        # Outputs a script framework to the console
        New-PSScriptScaffolding

        .EXAMPLE
        # Outputs a script framework to the console with a custon "Author Field"
        New-PSScriptScaffolding -Author "John Smith"

        .EXAMPLE
        # Outputs a script framework to a new script file
        New-PSScriptScaffolding -OutputPath "C:\Path\To\MyNewScript.ps1"

        .EXAMPLE
        # Overwrites a script framework to an existing script file
        New-PSScriptScaffolding -OutputPath "C:\Path\To\MyExistingScript.ps1"

        .EXAMPLE
        # Overwrites a script framework to an existing script file
        Get-Item "C:\Path\To\MyExistingScript.ps1" | New-PSScriptScaffolding

        .NOTES
        - PowerShell script files with the .ps1 extension are the only filetype supported. This is by design; .psm1 files should have an accompanying module manifest that contains the relevant metadata.

    #>

    [CmdletBinding(DefaultParameterSetName="Description")]
    PARAM
    (
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$Author,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [String]$CompanyName = $Author,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [String]$Version = '1.0.0',

        [Parameter(Mandatory=$false, Position=1, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("Filename","FullName")]
        [String]$OutputPath
    )

    # Create dynamic parameters
    DynamicParam
    {
        IF ($OutputPath)
        {
            # Create runtine dictionary
            $ParamDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

            # Define parameter attributes
            $Attribute  = [System.Management.Automation.ParameterAttribute]@{
                Mandatory = $false
            }

            # Create attribute collection
            $AttributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]$Attribute

            # Create parameter object and add name, type, attribute collection
            $PassThruParameter = [System.Management.Automation.RuntimeDefinedParameter]::new('PassThru',[Switch],$AttributeCollection)
            $PassThruParameter.Value = $false
            $PSBoundParameters['PassThru'] = $PassThruParameter.Value

            # Add parameter to runtime dictionary
            $ParamDictionary.Add('PassThru',$PassThruParameter)

            # Create parameter object and add name, type, attribute collection
            $ForceParameter = [System.Management.Automation.RuntimeDefinedParameter]::new('Force',[Switch],$AttributeCollection)
            $ForceParameter.Value = $false
            $PSBoundParameters['Force'] = $ForceParameter.Value

            # Add parameter to runtime dictionary
            $ParamDictionary.Add('Force',$ForceParameter)

            return $ParamDictionary
        }
    }

    BEGIN
    {
        # Transfer Dynamic Parameters to runtime
        $PassThru   = $PSBoundParameters['PassThru']
        $Force      = $PSBoundParameters['Force']

        # Locally scope ErrorActionPreference for predictable behavior of Try/Catch blocks inside the function
        $ErrorActionPreference = 'Stop'

        # Declare Script block
        $ScriptBlock = Get-Content -Path "$PSScriptRoot/../Templates/Template_Script.ps1" -Raw

        # Build parameter block for New-ScriptFileInfo command in PROCESS block
        $Params = @{
            Version     = $Version
            Author      = $Author
            CompanyName = $CompanyName
            Copyright   = "(c) $($(Get-Date).Year) $CompanyName"
            Description = "Blank"
        }
    }

    PROCESS
    {
        # Create output variable
        $Results = ''

        #region Generate Script Info
        $ScriptInfo = ((New-ScriptFileInfo @Params -PassThru).Split('>')[0]+'>' + "`r`n`r`n").TrimStart("`r`n")

        # Add ScriptInfo and ScriptBlock to results
        $Results = $ScriptInfo + $ScriptBlock

        IF ($OutputPath)
        {
            # Test whether output path is a file or directory
            SWITCH -Regex ($OutputPath)
            {
                # Test whether $OutputPath is a .ps1 file
                ".ps1$"
                {
                    $OutputFilename  = $OutputPath.Split('\')[-1]
                    $OutputDirectory = $OutputPath.Substring(0,$OutputPath.LastIndexOf('\'))
                }
                Default
                {
                    Write-Error "$OutputPath is not a .ps1 file. Please specify a valid filepath and try again." -ErrorAction Continue
                    Continue
                }
            }

            # Create directory if it doesn't exist
            IF (!(Test-Path $OutputDirectory))
            {
                TRY
                {
                    New-Item -Path $OutputDirectory -ItemType File -Force | Out-Null
                }
                CATCH
                {
                    Write-Warning "Unable to create directory $OutputDirectory. Defaulting to present working directory: $PWD"
                    $OutputDirectory = $PWD
                }
            }

            # Output file
            $Results | Out-File $OutputDirectory\$OutputFilename -Encoding ascii -Force:$Force
        }

        # Return results to the console if an output has not been specified or -Passthru has been selected
        IF(!$OutputPath -or $Passthru)
        {
            Write-Output $Results
        }
    }
    END
    {
        # Nothing to see here. Move along.
    }
}
