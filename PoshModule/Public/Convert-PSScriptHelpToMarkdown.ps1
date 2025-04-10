Function Convert-PSScriptHelpToMarkdown
{
    <#
        .SYNOPSIS
        Converts PowerShell script help to Markdown documentation.

        .DESCRIPTION
        Converts PowerShell script help to Markdown documentation. This function is capable of either outputting data to console, file, or both simultaneously.

        By default the function will capture help metadata at a script-level, and will not capture any function-level data. This functionality can be controlled with the `-IncludeNestedFunctions` parameter. For capturing documentation on files that only contain function data, please use `Convert-PSFunctionHelpToMarkdown` instead.

        The function will generate warnings for any missing data fields in order to notify the operator and ensure completeness of documentation.

        .PARAMETER Filename
        The name of the file to convert.

        .PARAMETER OutputPath
        The path to which to export the Markdown document.

        .PARAMETER HeaderGranularity
        Controls whether parameter names and examples are individually converted to Markdown headers. This setting will impact the appearance of the rendered text in addition to the metadata.

        .PARAMETER IncludeNestedFunctions
        Parses nested functions inside the file and converts the help for each to Markdown.

        .PARAMETER OutFile
        Controls whether an output file is generated.

        .PARAMETER PassThru
        Returns output to the console in addition to producing a file.

        .PARAMETER Force
        Forces an overwrite of an existing file.

        .PARAMETER Append
        Appends to an existing file. Necessary when consolidating documentation from multiple files into one.

        .EXAMPLE
        # Generate documentation to the console for a single file
        Convert-PSHelpToMarkdown -Filename C:\Path\To\Script.ps1

        .EXAMPLE
        # Generate documentation to the console for a single file, including any nested functions inside the file.
        Convert-PSHelpToMarkdown -Filename C:\Path\To\Script.ps1 -IncludeNestedFunctions

        .EXAMPLE
        # Generate documentation for a single file to a Markdown file with the same name and location, including any nested functions inside the file, and output to the console.
        Convert-PSHelpToMarkdown -Filename C:\Path\To\Script.ps1 -IncludeNestedFunctions -PassThru

        .EXAMPLE
        # Generate documentation for multiple files in a directory tree to a individual Markdown files with the same name and location as the .ps1 files
        Get-ChildItem C:\Path\To\Scripts*.ps1 -Recurse | Convert-PSHelpToMarkDown -IncludeNestedFunctions -OutFile

        .EXAMPLE
        # Generate documentation for multiple files in a directory tree to a individual Markdown files with the same name as the .ps1 files, but in a new folder
        Get-ChildItem C:\Path\To\Scripts*.ps1 -Recurse | Convert-PSHelpToMarkDown -IncludeNestedFunctions -OutFile -OutputPath C:\Path\To\Scripts\Docs

        .EXAMPLE
        # Generate documentation for multiple files in a directory tree to a single Markdown file
        Get-ChildItem C:\Path\To\Scripts*.ps1 -Recurse | Convert-PSHelpToMarkDown -IncludeNestedFunctions -OutFile -OutputPath C:\Path\To\Scripts\Consolidated.md -Force -Append

        .NOTES
        - Help data for dynamic parameters is not captured or converted. This is a known issue with Get-Help and not with this function. You can use examples that show dynamic parameters or move documentation of dynamic parameters to another section (such as the Description Field) in order to work around this limitation. [GitHub Issue](https://github.com/PowerShell/PowerShell/issues/6694)
        - This function does not support external help files at this time.

    #>

    [CmdletBinding(DefaultParameterSetName="__AllParameterSets")]
    PARAM
    (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("FullName","File")]
        [String]$Filename,

        [Parameter(Mandatory=$false, Position=1, ParameterSetName="OutFile")]
        [ValidateNotNullOrEmpty()]
        [String]$OutputPath,

        [Parameter(Mandatory=$false)]
        [ValidateSet("Coarse","Fine")]
        [String]$HeaderGranularity = "Fine",

        [Parameter(Mandatory=$false)]
        [Switch]$IncludeNestedFunctions,

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

        #region Script
        # Declare variables
        $FileTitle   = ''
        $FileHeader  = ''
        $Syntax      = ''
        $Synopsis    = ''
        $Description = ''
        $Parameters  = ''
        $Examples    = ''
        $Notes       = ''

        #region Script: Gather Info
        # Get script and function documentation
        $ScriptDocumentation   = Get-PSScriptDocumentation $Filename

        # Throw warning if function documentation is not found
        IF ($ScriptDocumentation -eq $null)
        {
            Write-Warning "Script documentation is null."
            Return
        }
        #endregion Script: Gather Info

        #region Script: File Info
        # Construct file title string
        $FileTitle = $ScriptDocumentation.Filename.Split('.')[0] | ConvertTo-MarkdownHeader

        # Add to results
        $Results.Add("$FileTitle`n") | Out-Null

        # Construct file header string
        $FileHeader += "FileName: $($ScriptDocumentation.Filename | Sort -Unique)`n"
        IF ($ScriptDocumentation.ScriptInfo.Version)
        {
            $FileHeader += "Version: $($ScriptDocumentation.ScriptInfo.Version)`n"
        }
        $FileHeader += "Generated on: $(Get-Date -Format MM.dd.yyyy)`n"

        # Convert file header to YAML code block
        $FileHeader = $FileHeader | Convertto-MarkdownCodeBlock -Language YAML

        # Add to results
        $Results.Add("$FileHeader`n") | Out-Null
        #endregion Script: File Info

        #region Script: Syntax
        # Construct syntax string
        $Syntax += "Syntax" | ConvertTo-MarkdownHeader -Level 2
        $Syntax += "`n"
        $Syntax += $ScriptDocumentation.Syntax.TrimEnd(' ') | Convertto-MarkdownCodeBlock -Language PowerShell
        $Syntax += "`n"

        # Add to results
        $Results.Add($Syntax) | Out-Null
        #endregion Script: Syntax

        #region Script: Synopsis
        # Construct synopsis string
        IF ($ScriptDocumentation.Synopsis -notmatch $ScriptDocumentation.FileName)
        {
            $Synopsis += "Synopsis" | ConvertTo-MarkdownHeader -Level 2
            $Synopsis += "`n$($ScriptDocumentation.Synopsis)`n"

            # Add to results
            $Results.Add($Synopsis) | Out-Null
        }
        #endregion Script: Synopsis

        #region Script: Description
        # Construct description string
        IF ($ScriptDocumentation.Description -notmatch $ScriptDocumentation.FileName)
        {
            $Description += "Description" | ConvertTo-MarkdownHeader -Level 2
            $Description += "`n$($ScriptDocumentation.Description.Text)`n"

            # Add to results
            $Results.Add($Description) | Out-Null
        }
        #endregion Script: Description

        #region Script: Parameters
        # Construct parameter string
        IF ($ScriptDocumentation.Parameters -ne '')
        {
            $Parameters += "Parameters" | ConvertTo-MarkdownHeader -Level 2
            $Parameters += "`n"

            FOREACH ($Parameter in $ScriptDocumentation.Parameters)
            {
                SWITCH($HeaderGranularity)
                {
                    "Coarse" {$Parameters += "-$($Parameter.Name)" | Convertto-MarkdownCodeBlock -Inline | ConvertTo-MarkdownFormattedText -Bold ; $Parameters += "`n"}
                    "Fine"   {$Parameters += "-$($Parameter.Name)" | Convertto-MarkdownCodeBlock -Inline | ConvertTo-MarkdownHeader -Level 3}
                }
                $Parameters += "`n"
                $Parameters += $Parameter.Description.Text
                $Parameters += "`n"
                $Parameters += "`n"
            }

            SWITCH($HeaderGranularity)
            {
                "Coarse" {$Parameters += "Parameter Details" | ConvertTo-MarkdownFormattedText -Bold}
                "Fine"     {$Parameters += "Parameter Details" | ConvertTo-MarkdownHeader -Level 3}
            }
            $Parameters += "`n"
            $Parameters += $ScriptDocumentation.Parameters | Select Name,DefaultValue,Required,ParameterValue,Position,PipelineInput | ConvertTo-MarkdownTable

            # Add to results
            $Results.Add($Parameters) | Out-Null
        }
        #endregion Script: Parameters

        #region Script: Examples
        # Construct examples string
        IF ($ScriptDocumentation.Examples)
        {
            $Examples += "Examples" | ConvertTo-MarkdownHeader -Level 2
            $Examples += "`n"
            FOREACH ($Example in $ScriptDocumentation.Examples)
            {
                SWITCH($HeaderGranularity)
                {
                    "Coarse" {$Examples += (Get-Culture).TextInfo.ToTitleCase($Example.Title.ToLower()) | ConvertTo-MarkdownFormattedText -Bold}
                    "Fine"   {$Examples += (Get-Culture).TextInfo.ToTitleCase($Example.Title.ToLower()) | ConvertTo-MarkdownHeader -Level 3}
                }
                $Examples += "`n"
                $Examples += $Example.Description.ToString() | Convertto-MarkdownCodeBlock -Language PowerShell
                $Examples += "`n"
                $Examples += "`n"
            }
            "$($ScriptDocumentation.Examples)`n"

            # Add to results
            $Results.Add($Examples) | Out-Null
        }
        #endregion Script: Examples

        #region Script: Notes
        # Construct notes string
        IF ($ScriptDocumentation.Notes -ne '')
        {
            $Notes += "Notes" | ConvertTo-MarkdownHeader -Level 2
            $Notes += "`n$($ScriptDocumentation.Notes)`n"

            # Add to results
            $Results.Add($Notes) | Out-Null
        }
        #endregion Script: Notes
        #endregion Script

        #region Functions
        IF ($IncludeNestedFunctions)
        {
            #region Functions: Gather Info
            # Get script and function documentation
            $FunctionDocumentation = Get-PSFunctionDocumentation $Filename

            # Throw warning if function documentation is not found
            IF ($FunctionDocumentation -eq $null)
            {
                Write-Warning "Function documentation is null."
                Return
            }
            #endregion Functions: Gather Info

            #region Functions: Header
            # Declare variables
            $FunctionHeader = ''

            # Construct functions header string
            $FunctionHeader += "Functions" | ConvertTo-MarkdownHeader -Level 2
            $FunctionHeader += "`n"

            # Add to results
            $Results.Add($FunctionHeader) | Out-Null
            #endregion Functions: Header

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
                $FunctionFileHeader += "$($Function.Name | Sort -Unique)`n" | ConvertTo-MarkdownHeader -Level 3

                # Add to results
                $Results.Add("$FunctionFileHeader`n") | Out-Null
                #endregion Functions: File Info

                #region Functions: Syntax
                # Construct syntax string
                $FunctionSyntax += "Syntax" | ConvertTo-MarkdownHeader -Level 4
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
                    $FunctionSynopsis += "Synopsis" | ConvertTo-MarkdownHeader -Level 4
                    $FunctionSynopsis += "`n$($Function.Synopsis)`n"

                    # Add to results
                    $Results.Add($FunctionSynopsis) | Out-Null
                }
                #endregion Functions: Synopsis

                #region Functions: Description
                # Construct description string
                IF ($Function.Description)
                {
                    $FunctionDescription += "Description" | ConvertTo-MarkdownHeader -Level 4
                    $FunctionDescription += "`n$($Function.Description.Text)`n"

                    # Add to results
                    $Results.Add($FunctionDescription) | Out-Null
                }
                #endregion Functions: Description

                #region Functions: Parameters
                # Construct parameter string
                IF ($Function.Parameters)
                {
                    $FunctionParameters += "Parameters" | ConvertTo-MarkdownHeader -Level 4
                    $FunctionParameters += "`n"

                    FOREACH ($Parameter in $Function.Parameters)
                    {
                        SWITCH($HeaderGranularity)
                        {
                            "Coarse" {$FunctionParameters += "-$($Parameter.Name)" | Convertto-MarkdownCodeBlock -Inline | ConvertTo-MarkdownFormattedText -Bold ; $FunctionParameters += "`n"}
                            "Fine"   {$FunctionParameters += "-$($Parameter.Name)" | Convertto-MarkdownCodeBlock -Inline | ConvertTo-MarkdownHeader -Level 5}
                        }
                        $FunctionParameters += "`n"
                        $FunctionParameters += $Parameter.Description.Text
                        $FunctionParameters += "`n"
                        $FunctionParameters += "`n"
                    }

                    SWITCH($HeaderGranularity)
                    {
                        "Coarse" {$FunctionParameters += "Parameter Details" | ConvertTo-MarkdownFormattedText -Bold}
                        "Fine"   {$FunctionParameters += "Parameter Details" | ConvertTo-MarkdownHeader -Level 5}
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
                    $FunctionExamples += "Examples" | ConvertTo-MarkdownHeader -Level 4
                    $FunctionExamples += "`n"
                    FOREACH ($Example in $Function.Examples)
                    {
                        SWITCH($HeaderGranularity)
                        {
                            "Coarse" {$FunctionExamples += (Get-Culture).TextInfo.ToTitleCase($Example.Title.ToLower()) | ConvertTo-MarkdownFormattedText -Bold}
                            "Fine"   {$FunctionExamples += (Get-Culture).TextInfo.ToTitleCase($Example.Title.ToLower()) | ConvertTo-MarkdownHeader -Level 5}
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
                    $FunctionNotes += "Notes" | ConvertTo-MarkdownHeader -Level 4
                    $FunctionNotes += "`n$($Function.Notes)`n"

                    # Add to results
                    $Results.Add($FunctionNotes) | Out-Null
                }
                #endregion Functions: Notes
            }
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
                        $OutputFilename  = "$($ScriptDocumentation.Filename.Split('.')[0]).md"
                        $OutputDirectory = $OutputPath
                    }
                    #endregion Output Handling: Folder
                }
            }
            ELSE
            #region Output Handling: Unspecified
            {
                $OutputFilename  = "$($Filename.Split('\')[-1].Split('.')[0]).md"
                $OutputDirectory = $Filename.Substring(0,$Filename.LastIndexOf('\'))
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