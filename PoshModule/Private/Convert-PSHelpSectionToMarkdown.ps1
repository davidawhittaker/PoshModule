Function Convert-PSHelpSectionToMarkdown
{
    <#
    .SYNOPSIS
    Converts .ps1 help headers to markdown format.

    .DESCRIPTION
    Converts headings in a .ps1 help file (such as '.PARAMETER') to markdown format.

    Any repeating heading (such as .PARAMETER) will be converted to subheadings ('##' in Markdown) and a first-level heading ('#' in Markdown) will be created above the first instance.

    If a single instance of a heading is found, it will be converted to a first-level heading ('#' in Markdown).

    .PARAMETER SearchTerm
    The name of the heading in the PowerShell heading to replace.

    .PARAMETER ReplacementTerm
    The name of the heading to be printed in Markdown. This will replace the value definied in SearchTerm.

    .PARAMETER Heading
    The name of the first-level Markdown heading to create when converting multiple instances of a PowerShell heading.

    .PARAMETER InputText
    The complete contents of the help file without the beginning and ending comment tags.

    .EXAMPLE
    $Contents     = Get-Content C:\Path\To\MyFunction.ps1
    $StartIndex   = ($Contents | Select-String "<#").LineNumber[0]
    $EndIndex     = ($Contents | Select-String "#`>").LineNumber[0] -2
    $HelpContents = $Contents[$StartIndex..$EndIndex]

    $HelpContents = Convert-PSHelpSectionToMarkdown -SearchTerm SYNOPSIS      -ReplacementTerm Synopsis    -Heading ''         -InputText $HelpContents
    $HelpContents = Convert-PSHelpSectionToMarkdown -SearchTerm DESCRIPTION   -ReplacementTerm Description -Heading ''         -InputText $HelpContents
    $HelpContents = Convert-PSHelpSectionToMarkdown -SearchTerm NOTES         -ReplacementTerm Notes       -Heading ''         -InputText $HelpContents
    $HelpContents = Convert-PSHelpSectionToMarkdown -SearchTerm PARAMETER     -ReplacementTerm ''          -Heading Parameters -InputText $HelpContents
    $HelpContents = Convert-PSHelpSectionToMarkdown -SearchTerm EXAMPLE       -ReplacementTerm ''          -Heading Examples   -InputText $HelpContents

    .NOTES
    - Tested on Windows on PowerShell 5.1.
    - Values for SearchTerms have been updated through PowerShell 7.4 LTS.
    - This is meant for internal use only and is intended as a helper function to be called in other Markdown conversion functions.
    - This tool is not an interpretive parser; it only searches strings for certain regex values and should be treated accordingly.
    - Nested comment blocks in the help file have the potential to break this function.
    - PSSCriptInfo blocks have the potential to break this function.

    #>

    PARAM
    (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateSet(
            "SYNOPSIS",
            "DESCRIPTION",
            "PARAMETER",
            "EXAMPLE",
            "INPUTS",
            "OUTPUTS",
            "NOTES",
            "LINK",
            "COMPONENT",
            "ROLE",
            "FUNCTIONALITY",
            "FORWARDHELPTARGETNAME",
            "FORWARDHELPCATEGORY",
            "REMOTEHELPRUNSPACE",
            "EXTERNALHELP"
        )]
        [ValidateNotNullOrEmpty()]
        [String]$SearchTerm,

        [Parameter(Mandatory=$true, Position=1)]
        [AllowEmptyString()]
        [String]$ReplacementTerm,

        [Parameter(Mandatory=$true, Position=2)]
        [AllowEmptyString()]
        [String]$Heading,

        [Parameter(Mandatory=$true, Position=3)]
        $InputText
    )

    # Create output variable
    $OutputText = $InputText

    # Split string if passed as a single text block
    IF ($OutputText.Count -eq 1)
    {
        $OutputText = $OutputText.Split("`n")
    }

    # Trim preceding whitespace
    $OutputText = $OutputText | %{$_.TrimStart(" ")}
    $OutputText = $OutputText | %{$_.TrimStart("`t")}

    # Get number of instances of search term
    $InstanceCount = ($OutputText | ?{$_ -match "^\.$SearchTerm"}).Count

    # Create a section header if 'InstanceCount' is greater than one and convert each instance to a subheading
    IF ($InstanceCount -gt 1)
    {
        # Set Counter to 0
        $i = 0

        WHILE ($i -lt $InstanceCount)
        {
            # Get index number for next occurence of search term
            # $IndexNumber = $OutputText.IndexOf($SearchTerm)
            $IndexNumber = ($OutputText | Select-String "^\.$SearchTerm").LineNumber[0] -1

            # Insert markdown heading before the first occurence of search term
            IF ($i -eq 0)
            {
                ($OutputText)[$IndexNumber] = ($OutputText)[$IndexNumber] | %{$_ -replace "^\.$SearchTerm","# $Heading`n`n## $ReplacementTerm"}
            }

            # Replace search term with markdown formatted replacement term
            ($OutputText)[$IndexNumber] = ($OutputText)[$IndexNumber] | %{$_ -replace "^\.$SearchTerm","## $ReplacementTerm"}

            # Fix a bug where if 'ReplacementTerm' is blank, the subheading will contain two spaces instead of one.
            ($OutputText)[$IndexNumber] = ($OutputText)[$IndexNumber] | %{$_ -replace "##  ","## "}

            # Increment counter
            $i++
        }
    }

    # Don't create a section header if 'InstanceCount' is equal to one; just replace the 'SearchTerm' with the 'ReplacementTerm'
    ELSE
    {
        # Get index number for next occurence of search term
        $IndexNumber = ($OutputText | Select-String "^\.$SearchTerm").LineNumber[0] -1

        # Replace search term with markdown formatted replacement term
        ($OutputText)[$IndexNumber] = ($OutputText)[$IndexNumber] | %{$_ -replace "^\.$SearchTerm","# $ReplacementTerm"}
    }

    return $OutputText
}