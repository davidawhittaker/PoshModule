{
    <#
        .SYNOPSIS


        .DESCRIPTION


        .PARAMETER


        .EXAMPLE
        # Description
        Command

        .NOTES
        - Lorem ipsum

    #>

    [CmdletBinding()]
    PARAM
    (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("MyAlias")]
        [String]$Parameter
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
        # Declare variables

    }
    #endregion PROCESS Block

    #region END Block
    END
    {
        Return $Results
    }
    #endregion END Block
}