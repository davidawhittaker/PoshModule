Function Get-PSFunctionDocumentation
{
    <#
        .SYNOPSIS
        Converts .ps1 help into a PowerShell object.

        .DESCRIPTION
        Parses .ps1 files using Abstract Syntax Tree (AST) to generate a standardized object format for function help files.

        The tool is compatible with multiple functions inside the same script file.

        .PARAMETER InputObject
        The path to the file to be analyzed. Relative paths are supported.

        Object can be also piped from Get-Command. The function must be loaded into your active PowerShell session for this to function properly.

        You can also simply pass name the function. The function must be loaded into your active PowerShell session for this to function properly.

        .EXAMPLE
        Get-PSFunctionDocumentation -File C:\Path\To\Function.ps1

        .EXAMPLE
        Get-PSFunctionDocumentation C:\Path\To\Function.ps1

        .EXAMPLE
        Get-PSFunctionDocumentation MyFunction

        .EXAMPLE
        Get-Command MyFunction | Get-PSFunctionDocumentation

        .EXAMPLE
        Get-ChildItem *.ps1 -Path C:\Path\To\Module -Recurse | Get-PSFunctionDocumentation

        .NOTES
        - This function is explicitly made for documenting functions and does not capture script-level parameters or help.
        - This function is designed to be a helper function in a larger tool.
        - This function does not support external help files at this time.

    #>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        $InputObject
    )

    #region BEGIN Block
    BEGIN
    {
        # Locally scope ErrorActionPreference for predictable behavior of Try/Catch blocks inside the function
        $ErrorActionPreference = 'Stop'

        # Create output variable
        $Results = [System.Collections.ArrayList]::new()
    }
    #endregion BEGIN Block

    #region PROCESS Block
    PROCESS
    {
        #region Input Handling
        # Execute search based on type of inputobject received

        SWITCH($InputObject.GetType().FullName)
        {
            "System.Management.Automation.FunctionInfo"
            {
                # Object was input from Get-Command ; transform it to a full definition and see if it can be found in PowerShell
                $Filename = $InputObject.Name
                $DefinitionText = "Function $Filename {`n$(Get-Content Function:$Filename)`n}"

                # Query AST
                $AST = [System.Management.Automation.Language.Parser]::ParseInput($DefinitionText,[ref]$null,[ref]$null)
            }

            "System.IO.FileInfo"
            {
                # Check if object is a path
                $Filename = (Resolve-Path $InputObject.Fullname).Path

                # Query AST
                $AST = [System.Management.Automation.Language.Parser]::ParseFile($Filename,[ref]$null, [ref]$null)
            }

            "System.String"
            {
                TRY
                {
                    # Check if object is a path
                    $Filename = (Resolve-Path $InputObject).Path

                    # Query AST
                    $AST = [System.Management.Automation.Language.Parser]::ParseFile($Filename,[ref]$null, [ref]$null)
                }
                CATCH
                {
                    # Object is likely to be a function name ; transform it to a full definition and see if it can be found in PowerShell
                    $Filename = $InputObject
                    $DefinitionText = "Function $Filename {`n$(Get-Content Function:$Filename)`n}"

                    # Query AST
                    $AST = [System.Management.Automation.Language.Parser]::ParseInput($DefinitionText,[ref]$null,[ref]$null)
                }
            }
        }
        #endregion Input Handling

        # $Filename

        # Declare variables
        $Functions     = [System.Collections.ArrayList]::new()
        $FilenameShort = $Filename.Split('\')[-1]

        # Construct predicate to execute search
        $Predicate = {
            PARAM([System.Management.Automation.Language.Ast] $Ast)
            $AST -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                (
                    $PSVersionTable.PSVersion.Major -lt 5 -or
                    $Ast.Parent -isnot [System.Management.Automation.Language.FunctionMemberAst]
                )
        }

        # Execute AST search
        $FunctionDefinitions = $Ast.FindAll($Predicate, $true)

        FOREACH ($FunctionDefinition in $FunctionDefinitions)
        {
            Write-Verbose "Processing Function $($FunctionDefinition.Name)"

            # Declare variables
            $Syntax      = ''
            $Synopsis    = ''
            $Description = ''
            $Parameters  = ''
            $Notes       = ''
            $Examples    = [System.Collections.ArrayList]::new()

            #region Parameters and Help Block
            TRY
            {
                # Declare variables
                $ParamBlock   = ''
                $CommentBlock = ''
                $HelpFile     = ''

                # Look up existing PARAM block
                $ParamBlock = $FunctionDefinition.Body.ParamBlock.Extent.Text | Out-String

                # Look up existing help block
                IF ($FunctionDefinition.GetHelpContent())
                {
                    $CommentBlock = $FunctionDefinition.GetHelpContent().GetCommentBlock() | Out-String
                }

                # Create an ephemeral function that only contains the help file and parameters
                $ScriptBlock = [scriptblock]::Create(('
                Function {0} {{
                    {1}
                    {2}
                }}' -f $FunctionDefinition.Name, $CommentBlock, $ParamBlock))

                # Dot-source the ephemeral function to pull help file into an acceptable PowerShell object format
                $HelpFile = & {. $ScriptBlock ; Get-Help $FunctionDefinition.Name}
            }

            CATCH
            {
                Write-Warning "$($FilenameShort): Unable to generate comment block / help file for $($FunctionDefinition.Name)."
            }
            #endregion Parameters and Help Block

            #region Syntax
            # Dot-source the ephemeral function to query command syntax
            $Syntax = & {. $ScriptBlock ; (Get-Command $FunctionDefinition.Name -Syntax -ErrorAction SilentlyContinue)}

            # Trim carriage returns
            $Syntax = $Syntax.Split("`r`n") | ?{$_ -ne ""}
            $Syntax = $Syntax -Join "`r`n"

            $EscapedSyntax = $Syntax -replace "\[", "``[" -replace "\]", "``]"
            #endregion Syntax

            #region Synopsis
            # Trim carriage returns
            $Synopsis = $HelpFile.Synopsis.Replace("`n", "").Replace("`r", "")

            # Blank Synopsis field if it is populated by auto-generated syntax
            IF ($Synopsis -like "*$EscapedSyntax*")
            {
                Write-Warning "$($FilenameShort): Field for $($FunctionDefinition.Name) is blank: Synopsis."
                $Synopsis = ""
            }
            ELSE
            {
                $Synopsis = $HelpFile.Synopsis
            }
            #endregion Synopsis

            #region Description
            IF (!$HelpFile.Description)
            {
                Write-Warning "$($FilenameShort): Field is blank: Description. Using synopsis field to auto-populate if it exists."
                $Description = $HelpFile.Synopsis
            }
            ELSE
            {
                $Description = $HelpFile.Description
            }
            #endregion Description

            #region Parameters
            IF(!$HelpFile.Parameters.Parameter.Description)
            {
                Write-Warning "$($FilenameShort): Field for $($FunctionDefinition.Name) is blank: Parameters."
                $Parameters = ''
            }
            ELSE
            {
                $Parameters = $HelpFile.Parameters.Parameter
            }
            #endregion Parameters

            #region Examples
            IF (!$HelpFile.Examples -or !$HelpFile.Examples.Example.Description)
            {
                Write-Warning "$($FilenameShort): Field for $($FunctionDefinition.Name) is blank: Examples."
                $Examples = ''
            }
            ELSE
            {
                # Add properly formatted examples to $Examples object
                FOREACH ($Entry in $HelpFile.Examples.Example)
                {
                    # Add the rest of the help object if there's a multi-line example
                    IF ($Entry.Remarks.Text -match "\S")
                    {
                        $Example = [PSCustomObject]@{
                            Title        = $Entry.Title.TrimStart('-').TrimEnd('-').Trim(' ')
                            Description  = $Entry.Code + "`n" + $Entry.Remarks.Text
                        }
                    }
                    ELSE
                    {
                        $Example = [PSCustomObject]@{
                            Title        = $Entry.Title.TrimStart('-').TrimEnd('-').Trim(' ')
                            Description  = $Entry.Code
                        }
                    }
                    $Examples.Add($Example) | Out-Null
                }
            }
            #endregion Examples

            #region Notes
            IF(!$HelpFile.AlertSet)
            {
                Write-Warning "$($FilenameShort): Field for $($FunctionDefinition.Name) is blank: Notes."
                $Notes = ''
            }
            ELSE
            {
                $Notes = $HelpFile.AlertSet.Alert.Text
            }
            #endregion Notes

            # Construct function object and add to larger functions object
            $Function = [PSCustomObject]@{
                Name        = $FunctionDefinition.Name
                FileName    = $Filename
                Synopsis    = $Synopsis
                Description = $Description
                Parameters  = $Parameters
                Syntax      = $Syntax
                Examples    = $Examples
                Notes       = $Notes
            }

            $Functions.Add($Function) | Out-Null
        }

        $Results.Add($Functions) | Out-Null
    }
    #endregion PROCESS Block

    #region END Block
    END
    {
        Return $Results
    }
    #endregion END Block
}