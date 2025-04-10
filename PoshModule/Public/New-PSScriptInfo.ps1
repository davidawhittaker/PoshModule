Function New-PSScriptInfo
{
    <#
        .SYNOPSIS
        Generates PSScriptFileInfo metadata.

        .DESCRIPTION
        Generates PSScriptFileInfo metadata to the console, a new script file, or an existing script file.

        .PARAMETER Author
        The name of the script's author.

        .PARAMETER CompanyName
        The company name of the script's author.

        .PARAMETER Description
        A detailed description of the script's intended purpose.

        .PARAMETER SkipDescription
        Skips creation of the `Description` field in the script. This can be useful if a full help file already exists or is to be implemented.

        .PARAMETER OutputPath
        The path to which to export the script. Absolute and relative paths are both supported.

        .PARAMETER Force
        Forces an overwrite of an existing file.

        .PARAMETER PassThru
        Returns output to the console in addition to producing a file.

        .EXAMPLE
        # Outputs a script info block to the console
        New-PSScriptInfo -Description "My description."

        .EXAMPLE
        # Outputs a script info block to the console, omitting the description field
        New-PSScriptInfo -SkipDescription

        .EXAMPLE
        # Outputs a script info block to the console with a custon "Author Field"
        New-PSScriptInfo -Description "My description." -Author "John Smith"

        .EXAMPLE
        # Outputs a script info block to a new script file
        New-PSScriptInfo -Description "My description." -OutputPath "C:\Path\To\MyNewScript.ps1"

        .EXAMPLE
        # Outputs a script info block to an existing script file
        New-PSScriptInfo -Description "My description." -OutputPath "C:\Path\To\MyExistingScript.ps1"

        .EXAMPLE
        # Outputs a script info block to an existing script file
        Get-Item "C:\Path\To\MyExistingScript.ps1" | New-PSScriptInfo -SkipDescription

        .EXAMPLE
        # Outputs a script info block to all existing script files in a folder; each will contain a unique GUID
        Get-ChildItem "C:\Path\To\MyExistingScripts\*.ps1" -File | New-PSScriptInfo -SkipDescription

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
        [String]$Version = "1.0.0",

        [Parameter(Mandatory=$true, Position=0, ParameterSetName="Description")]
        [ValidateNotNullOrEmpty()]
        [String]$Description,

        [Parameter(Mandatory=$false, ParameterSetName="SkipDescription")]
        [Switch]$SkipDescription,

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

        # Determine Description based on Parameters supplied by the user
        $ScriptDescription = SWITCH ($SkipDescription)
        {
            $true  {"Blank"}
            $false {$Description}
        }

        # Build parameter block for New-ScriptFileInfo command in PROCESS block
        $Params = @{
            Version     = "1.0.0"
            Author      = $Author
            CompanyName = $CompanyName
            Copyright   = "(c) $($(Get-Date).Year) $CompanyName"
            Description = $ScriptDescription
        }
    }

    PROCESS
    {
        # Create output variable
        $Results = ''

        #region Generate Script Info
        $ScriptInfo = SWITCH ($SkipDescription)
        {
            $true  {(New-ScriptFileInfo @Params -PassThru).Split('>')[0]+'>' + "`r`n`r`n"}
            $false {(New-ScriptFileInfo @Params -PassThru).Replace('Param()','')}
        }

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

            # Try to get file contents ; on failure create file if it doesn't exist
            TRY
            {
                $FileContents = Get-Content $OutputPath -Raw
            }
            CATCH
            {
                TRY
                {
                    Resolve-Path $OutputPath | Out-Null

                    # If file exists and function is unable to retrieve contents, throw error
                    Write-Error "Unable to retrieve contents for file $OutputPath. " -ErrorAction Continue
                    Continue
                }
                CATCH
                {
                    New-Item $OutputPath -ItemType File | Out-Null

                    $FileContents = ''
                }
            }

            # Add ScriptInfo to beginning of file contents
            $Results = $ScriptInfo + $FileContents

            # Output file
            $Results | Out-File $OutputDirectory\$OutputFilename -Encoding ascii -Force:$Force
        }
        ELSE
        {
            $Results = $ScriptInfo
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
