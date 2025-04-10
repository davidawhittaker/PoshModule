Function Convert-PSFunctionHelpToMarkdown
{
    <#
        .SYNOPSIS
        Converts PowerShell script help to Markdown documentation.

        .DESCRIPTION
        Converts PowerShell script help to Markdown documentation. This function is capable of either outputting data to console, file, or both simultaneously.

        By default the function will capture help metadata at a function-level, and will not capture any script-level data. This functionality can be controlled with the `-IncludeNestedFunctions` parameter. For capturing documentation in script files, please use `Convert-PSScriptHelpToMarkdown` instead.

        The function will generate warnings for any missing data fields in order to notify the operator and ensure completeness of documentation.

        .PARAMETER InputObject
        The name of the file to convert. Fully qualified and relative paths are both supported.

        If functions have been loaded into memory via dot-sourcing or module import, they may be passed via Get-Command or by name.

        .PARAMETER OutputPath
        The path to which to export the Markdown document. Folder paths and individual filenames are both supported.

        .PARAMETER HeaderGranularity
        Controls whether parameter names and examples are individually converted to Markdown headers. This setting will impact the appearance of the rendered text in addition to the metadata.

        .PARAMETER OutFile
        Controls whether an output file is generated.

        If used in conjunction with `-OutputPath`, a file can be generated in the desired destination.

        If used without `-OutputPath`, the files will be generated in the same directory in which the script files reside. If functions are ready from memory instead of from a file, the destination file will default to the present working directory of the active PowerShell session.

        .PARAMETER PassThru
        Returns output to the console in addition to producing a file.

        .PARAMETER Force
        Forces an overwrite of an existing file.

        .PARAMETER Append
        Appends to an existing file. Necessary when consolidating documentation from multiple files into one.

        .EXAMPLE
        # Generate documentation to the console for a single file
        Convert-PSFunctionHelpToMarkdown C:\Path\To\Script.ps1

        .EXAMPLE
        # Generate documentation to the console for a function that has been loaded into memory.
        Convert-PSFunctionHelpToMarkdown MyFunction

        .EXAMPLE
        # Generate documentation for a single file to a Markdown file with the same name and location, including any nested functions inside the file, and output to the console.
        Convert-PSFunctionHelpToMarkdown C:\Path\To\Script.ps1 -PassThru

        .EXAMPLE
        # Generate documentation for multiple files in a directory tree to individual Markdown files with the same name and location as the .ps1 files
        Get-ChildItem C:\Path\To\Functions\*.ps1 -Recurse | Convert-PSFunctionHelpToMarkdown -OutFile

        .EXAMPLE
        # Generate documentation for multiple files in a directory tree to individual Markdown files with the same name as the .ps1 files, but in a new folder
        Get-ChildItem C:\Path\To\Functions\*.ps1 -Recurse | Convert-PSFunctionHelpToMarkdown -OutFile -OutputPath C:\Path\To\Functions\Docs

        .EXAMPLE
        # Generate documentation for multiple files in a directory tree to a single Markdown file
        Get-ChildItem C:\Path\To\Functions\*.ps1 -Recurse | Convert-PSFunctionHelpToMarkdown -OutFile -OutputPath C:\Path\To\Functions\Consolidated.md -Force -Append

        .EXAMPLE
        # Generate documentation for a module to individual Markdown files with the same name as each function
        Get-Command -Module MyModule -CommandType Function | Convert-PSFunctionHelpToMarkdown -OutFile -OutputPath C:\Path\To\Docs

        .NOTES
        - Help data for dynamic parameters is not captured or converted. This is a known issue with Get-Help and not with this function. You can use examples that show dynamic parameters or move documentation of dynamic parameters to another section (such as the Description Field) in order to work around this limitation. [GitHub Issue](https://github.com/PowerShell/PowerShell/issues/6694)
        - This function does not support external help files at this time.

    #>

    [CmdletBinding(DefaultParameterSetName="__AllParameterSets")]
    PARAM
    (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("FullName")]
        $InputObject,

        [Parameter(Mandatory=$false, Position=1, ParameterSetName="OutFile")]
        [ValidateNotNullOrEmpty()]
        [String]$OutputPath,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Coarse","Fine")]
        [String]$HeaderGranularity = "Fine",

        [Parameter(Mandatory=$false, ParameterSetName="OutFile")]
        [Switch]$OutFile
    )

    # Create dynamic parameters
    DynamicParam
    {
        IF ($PSCmdlet.ParameterSetName -eq "OutFile")
        {
            # Create runtine dictionary
            $ParamDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

            # Define parameter attributes
            $Attribute  = [System.Management.Automation.ParameterAttribute]@{
                Mandatory        = $false
                ParameterSetName = "OutFile"
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

            # Create parameter object and add name, type, attribute collection
            $AppendParameter = [System.Management.Automation.RuntimeDefinedParameter]::new('Append',[Switch],$AttributeCollection)
            $AppendParameter.Value = $false
            $PSBoundParameters['Append'] = $AppendParameter.Value

            # Add parameter to runtime dictionary
            $ParamDictionary.Add('Append',$AppendParameter)

            return $ParamDictionary
        }
    }

    #region BEGIN Block
    BEGIN
    {
        # Transfer Dynamic Parameters to runtime
        $Append   = $PSBoundParameters['Append']
        $PassThru = $PSBoundParameters['PassThru']
        $Force    = $PSBoundParameters['Force']

        # Locally scope ErrorActionPreference for predictable behavior of Try/Catch blocks inside the function
        $ErrorActionPreference = 'Stop'
    }
    #endregion BEGIN Block

    #region PROCESS Block
    PROCESS
    {
        # Create output variable
        $Results = [System.Collections.ArrayList]::new()

        #region Functions
        #region Functions: Gather Info
        # Get script and function documentation
        $FunctionDocumentation = Get-PSFunctionDocumentation $InputObject

        # Throw warning if function documentation is not found
        IF ($FunctionDocumentation -eq $null)
        {
            Write-Warning "Function documentation is null."
            Return
        }
        #endregion Functions: Gather Info

        FOREACH ($Function in $FunctionDocumentation)
        {
            # Declare variables
            $FunctionFileHeader  = ''
            $FunctionSyntax      = ''
            $FunctionSynopsis    = ''
            $FunctionDescription = ''
            $FunctionParameters  = ''
            $FunctionExamples    = ''
            $FunctionNotes       = ''

            #region Functions: File Info
            # Construct file header string
            $FunctionFileHeader += "$($Function.Name | Sort -Unique)`n" | ConvertTo-MarkdownHeader

            # Add to results
            $Results.Add("$FunctionFileHeader`n") | Out-Null
            #endregion Functions: File Info

            #region Functions: Syntax
            # Construct syntax string
            $FunctionSyntax += "Syntax" | ConvertTo-MarkdownHeader -Level 2
            $FunctionSyntax += "`n"
            $FunctionSyntax += $Function.Syntax | Convertto-MarkdownCodeBlock -Language PowerShell
            $FunctionSyntax += "`n"

            # Add to results
            $Results.Add($FunctionSyntax) | Out-Null
            #endregion Functions: Syntax

            #region Functions: Synopsis
            # Construct synopsis string
            IF ($Function.Synopsis)
            {
                $FunctionSynopsis += "Synopsis" | ConvertTo-MarkdownHeader -Level 2
                $FunctionSynopsis += "`n$($Function.Synopsis)`n"

                # Add to results
                $Results.Add($FunctionSynopsis) | Out-Null
            }
            #endregion Functions: Synopsis

            #region Functions: Description
            # Construct description string
            IF ($Function.Description)
            {
                $FunctionDescription += "Description" | ConvertTo-MarkdownHeader -Level 2
                $FunctionDescription += "`n$($Function.Description.Text)`n"

                # Add to results
                $Results.Add($FunctionDescription) | Out-Null
            }
            #endregion Functions: Description

            #region Functions: Parameters
            # Construct parameter string
            IF ($Function.Parameters)
            {
                $FunctionParameters += "Parameters" | ConvertTo-MarkdownHeader -Level 2
                $FunctionParameters += "`n"

                FOREACH ($Parameter in $Function.Parameters)
                {
                    SWITCH($HeaderGranularity)
                    {
                        "Coarse" {$FunctionParameters += "-$($Parameter.Name)" | Convertto-MarkdownCodeBlock -Inline | ConvertTo-MarkdownFormattedText -Bold ; $FunctionParameters += "`n"}
                        "Fine"   {$FunctionParameters += "-$($Parameter.Name)" | Convertto-MarkdownCodeBlock -Inline | ConvertTo-MarkdownHeader -Level 3}
                    }
                    $FunctionParameters += "`n"
                    $FunctionParameters += $Parameter.Description.Text
                    $FunctionParameters += "`n"
                    $FunctionParameters += "`n"
                }

                SWITCH($HeaderGranularity)
                {
                    "Coarse" {$FunctionParameters += "Parameter Details" | ConvertTo-MarkdownFormattedText -Bold}
                    "Fine"   {$FunctionParameters += "Parameter Details" | ConvertTo-MarkdownHeader -Level 3}
                }
                $FunctionParameters += "`n"
                $FunctionParameters += $Function.Parameters | Select Name,DefaultValue,Required,ParameterValue,Position,PipelineInput | ConvertTo-MarkdownTable

                # Add to results
                $Results.Add($FunctionParameters) | Out-Null
            }
            #endregion Functions: Parameters

            #region Functions: Examples
            # Construct examples string
            IF ($Function.Examples)
            {
                $FunctionExamples += "Examples" | ConvertTo-MarkdownHeader -Level 2
                $FunctionExamples += "`n"
                FOREACH ($Example in $Function.Examples)
                {
                    SWITCH($HeaderGranularity)
                    {
                        "Coarse" {$FunctionExamples += (Get-Culture).TextInfo.ToTitleCase($Example.Title.ToLower()) | ConvertTo-MarkdownFormattedText -Bold}
                        "Fine"   {$FunctionExamples += (Get-Culture).TextInfo.ToTitleCase($Example.Title.ToLower()) | ConvertTo-MarkdownHeader -Level 3}
                    }
                    $FunctionExamples += "`n"
                    $FunctionExamples += $Example.Description.ToString() | Convertto-MarkdownCodeBlock -Language PowerShell
                    $FunctionExamples += "`n"
                    $FunctionExamples += "`n"
                }
                "$($Function.Examples)`n"

                # Add to results
                $Results.Add($FunctionExamples) | Out-Null
            }
            #endregion Functions: Examples

            #region Functions: Notes
            # Construct notes string
            IF ($Function.Notes)
            {
                $FunctionNotes += "Notes" | ConvertTo-MarkdownHeader -Level 2
                $FunctionNotes += "`n$($Function.Notes)`n"

                # Add to results
                $Results.Add($FunctionNotes) | Out-Null
            }
            #endregion Functions: Notes
        }
        #endregion Functions

        #region Output Handling
        IF ($PSCmdlet.ParameterSetName -eq "OutFile")
        {
            IF ($OutputPath)
            {
                # Test whether output path is a file or directory
                SWITCH -Regex ($OutputPath)
                {
                    #region Output Handling: File
                    # Absolute regex madness
                    "\.\w\w$|\.\w\w\w$|\.\w\w\w\w$|\.markdown$"
                    {
                        $OutputFilename  = $OutputPath.Split('\')[-1]
                        $OutputDirectory = $OutputPath.Substring(0,$OutputPath.LastIndexOf('\'))
                    }
                    #endregion Output Handling: File
                    #region Output Handling: Folder
                    Default
                    {
                        $OutputFilename  = "$($FunctionDocumentation.Name.Split('.')[0]).md"
                        $OutputDirectory = $OutputPath
                    }
                    #endregion Output Handling: Folder
                }
            }
            ELSE
            #region Output Handling: Unspecified
            {
                TRY
                {
                    $OutputFilename  = "$($FunctionDocumentation.Filename.Split('\')[-1].Split('.')[0]).md"
                    $OutputDirectory = $FunctionDocumentation.Filename.Substring(0,$FunctionDocumentation.Filename.LastIndexOf('\'))
                }
                CATCH
                {
                    Write-Warning "Unable to determine function directory. Defaulting to present working directory: $PWD"
                    $OutputFilename = "$($FunctionDocumentation.Filename).md"
                    $OutputDirectory = $PWD
                }
            }
            #endregion Output Handling: Unspecified

            # Create path if it doesn't exist
            IF (!(Test-Path $OutputDirectory))
            {
                New-Item -Path $OutputDirectory -ItemType Directory | Out-Null
            }

            # Output file
            $Results | Out-File $OutputDirectory\$OutputFilename -Encoding ascii -Force:$Force -Append:$Append
        }
        #endregion Output Handling

        #region Results
        # Return results to the console if an output has not been specified or -Passthru has been selected
        IF(($PSCmdlet.ParameterSetName -ne "OutFile") -or $Passthru)
        {
            Return $Results
        }
        #endregion Results
    }
    #endregion PROCESS Block

    #region END Block
    END
    {
        # Nothing to see here. Move along.
    }
    #endregion END Block
}