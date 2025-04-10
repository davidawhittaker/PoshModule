Function Get-PSScriptDocumentation
{
    <#
        .SYNOPSIS
        Converts .ps1 help into a PowerShell object.

        .DESCRIPTION
        Parses .ps1 files using Get-Help and Get-ScriptFileInfo to generate a standardized object format for script help files.

        .PARAMETER FileName
        The path to the file to be analyzed. Relative paths are supported.

        .EXAMPLE
        Get-PSScriptDocumentation -File C:\Path\To\Function.ps1

        .EXAMPLE
        Get-PSScriptDocumentation C:\Path\To\Function.ps1

        .EXAMPLE
        Get-ChildItem *.ps1 -Path C:\Path\To\Scripts -Recurse | Get-PSScriptDocumentation

        .NOTES
        - This function is explicitly made for documenting script-level parameters and help; it does not capture parameters or help for nested functions.
        - This function is designed to be a helper function in a larger tool.

    #>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("FullName")]
        [String]$Filename
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
        IF ($Filename -notmatch ".ps1$")
        {
            Write-Warning "File $Filename is not a PowerShell File. Skipping."
            return
        }

        # Declare variables
        $FilenameShort = $Filename.Split('\')[-1]
        $ScriptInfo    = ''
        $Synopsis      = ''
        $Syntax        = ''
        $Description   = ''
        $Parameters    = ''
        $Notes         = ''
        $Examples      = [System.Collections.ArrayList]::new()

        Write-Verbose "Processing File $($FilenameShort)"

        $ScriptInfo =
        TRY
        {
            Test-ScriptFileInfo $Filename
        }
        CATCH
        {
            Write-Warning "$($FilenameShort): Unable to generate PSScriptInfo."
        }

        # Get help file object for script
        $HelpFile = Get-Help -Full $Filename

        #region Syntax
        # Query command syntax for the script
        $Syntax = Get-Command $Filename -Syntax -ErrorAction SilentlyContinue

        # Trim carriage returns
        $Syntax = $Syntax.Split("`r`n") | ?{$_ -ne ""}
        $Syntax = $Syntax -Join "`r`n"
        #endregion Syntax

        #region Synopsis
        IF (!$HelpFile.Synopsis)
        {
            Write-Warning "$($FilenameShort): Field is blank: Synopsis."
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
        IF(!$HelpFile.Parameters)
        {
            Write-Warning "$($FilenameShort): Field is blank: Parameters."
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
            Write-Warning "$($FilenameShort): Field is blank: Examples."
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
                        Description  = $Entry.Code + "`n" + $Entry.Remarks.Text.Trim(' ')
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

        # Construct ScriptDetails object to add to Results
        $ScriptDetails = [PSCustomObject]@{
            FileName    = $FilenameShort
            ScriptInfo  = $ScriptInfo
            Synopsis    = $Synopsis
            Description = $Description
            Parameters  = $Parameters
            Syntax      = $Syntax
            Examples    = $Examples
            Notes       = $Notes
        }

        $Results.Add($ScriptDetails) | Out-Null
    }
    #endregion PROCESS Block

    #region END Block
    END
    {
        Return $Results
    }
    #endregion END Block
}