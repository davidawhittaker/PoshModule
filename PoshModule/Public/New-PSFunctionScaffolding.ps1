Function New-PSFunctionScaffolding
{
    <#
        .SYNOPSIS
        Creates a new PowerShell function.

        .DESCRIPTION
        Creates a new PowerShell function and outputs the results to either the console or a file.

        .PARAMETER Name
        The name of the function to be created. Names should adhere to PowerShell naming conventions.

        .PARAMETER OutputPath
        The path to which to export the function. Folder paths and individual filenames are both supported.

        .PARAMETER Module
        Creates a function file inside the directory structure for a given module. If the module is not found in PowerShell because it is not in $ENV:PSModulePath, it will default to the present working directory.

        This function is designed to work with modules that have the following directory structure:

        - Root
            - Manifest.psd1
            - Module.psm1
            - Public
                - PublicFunction1.ps1
                - PublicFunction2.ps1
            - Private
                - PrivateFunction1.ps1
                - PrivateFunction2.ps1

        By default, function files will be placed in the Public Directory. This behavior can be controlled with the `-Visibility` parameter.

        If you wish to place a new function inside an existing monolithic .psm1 or .ps1 file, please use the `-OutputPath` parameter.

        .PARAMETER PassThru
        Returns output to the console in addition to producing a file.

        .PARAMETER Force
        Forces an overwrite of an existing file.

        .PARAMETER Append
        Appends to an existing file. Necessary when consolidating documentation from multiple files into one.

        .PARAMETER Visibility
        Controls whether the function will be placed in the Public or Private sub-directories of a given module.

        .EXAMPLE
        # Outputs the framework for a new function to the console
        New-PSFunctionScaffolding -Name Test-Function

        .EXAMPLE
        # Outputs the framework for a new function to a file with the same name
        New-PSFunctionScaffolding -Name Test-Function -OutputPath C:\Path\To\Module\Public

        .EXAMPLE
        # Outputs the framework for a new function to a file with a unique name
        New-PSFunctionScaffolding -Name Test-Function -OutputPath C:\Path\To\Module\Public\MyFunction.ps1

        .EXAMPLE
        # Outputs the framework for a new function to an existing script
        New-PSFunctionScaffolding -Name Test-Function -OutputPath C:\Path\To\Scripts\MyExistingScript.ps1 -Append -Force

        .EXAMPLE
        # Outputs the framework for multiple new private functions to an existing module
        New-PSFunctionScaffolding -Name Test-Function,Test-Function2 -Module MyExistingModule -Visibility Private

        .NOTES
        - Use of improper naming conventions will cause this command to fail. This is by design and ensures proper naming standards are adhered to.

    #>

    [CmdletBinding(DefaultParameterSetName="__AllParameterSets")]
    PARAM
    (
        [Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [String[]]$Name,

        [Parameter(Mandatory=$false, Position=1, ParameterSetName="OutFile")]
        [ValidateNotNullOrEmpty()]
        [String]$OutputPath,

        [Parameter(Mandatory=$false, Position=1, ParameterSetName="Module")]
        [ValidateNotNullOrEmpty()]
        [String]$Module
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

        IF ($PSCmdlet.ParameterSetName -eq "Module")
        {
            # Create runtine dictionary
            $ParamDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

            # Define parameter attributes
            $Attribute  = [System.Management.Automation.ParameterAttribute]@{
                Mandatory        = $false
                ParameterSetName = "Module"
            }

            # Create attribute collection
            $AttributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]$Attribute

            # Add ValidateSet to attribute collection
            $AttributeCollection.Add([System.Management.Automation.ValidateSetAttribute]::new("Public","Private"))

            # Create parameter object and add name, type, attribute collection
            $VisibilityParameter = [System.Management.Automation.RuntimeDefinedParameter]::new('Visibility',[String],$AttributeCollection)
            $VisibilityParameter.Value = 'Public'
            $PSBoundParameters['Visibility'] = $VisibilityParameter.Value

            # Add parameter to runtime dictionary
            $ParamDictionary.Add('Visibility',$VisibilityParameter)

            return $ParamDictionary
        }
    }

    #region BEGIN Block
    BEGIN
    {
        # Transfer Dynamic Parameters to runtime
        $Append     = $PSBoundParameters['Append']
        $PassThru   = $PSBoundParameters['PassThru']
        $Force      = $PSBoundParameters['Force']
        $Visibility = $PSBoundParameters['Visibility']

        # Locally scope ErrorActionPreference for predictable behavior of Try/Catch blocks inside the function
        $ErrorActionPreference = 'Stop'

        # Declare variables
        $FunctionBlock = Get-Content -Path "$PSScriptRoot/../Templates/Template_Function.ps1" -Raw
    }
    #endregion BEGIN Block

    #region PROCESS Block
    PROCESS
    {
        # Create output variable
        $Results = [System.Collections.ArrayList]::new()

        FOREACH ($Function in $Name)
        {
            #region Input Validation: Name
            #region Input Validation: Name: Conventions
            # Validate that the input string matches the Verb-Noun naming convention
            IF ($Function -notmatch "(?-i:^\w+-\w+$)")
            {
                Write-Error "$Function does not match PowerShell Verb-Noun naming conventions. Please choose a more appropriate name." -ErrorAction Continue
                Continue
            }
            #endregion Input Validation: Name: Conventions

            #region Input Validation: Name: Approved Verbs
            # Get list of approved verbs
            $ApprovedVerbs = (Get-Verb).Verb

            # Split function name to get verb
            $Verb = $Function.Split('-')[0]

            IF ($ApprovedVerbs -notcontains $Verb)
            {
                Write-Error "$Function verb '$Verb' does not match PowerShell approved verbs. For a list of approved verbs, please run Get-Verb." -ErrorAction Continue
                Continue
            }

            #endregion Input Validation: Name: Approved Verbs

            #region Input Validation: Name: Capitalization
            # Validate that the input string has proper capitalization
            IF ($Function -notmatch "(?-i:^[A-Z]\w+-[A-Z]\w+$)")
            {
                # Declare variables
                $CapitalizedName = @()

                # Split input string into segments
                $Segments = $Function.Split('-')

                # Force title case for each segment
                FOREACH ($Segment in $Segments)
                {
                    # Convert to lowercase so titlecase function works properly
                    $Segment = $Segment.ToLower()

                    # Convert to titlecase
                    $CapitalizedName += (Get-Culture).TextInfo.ToTitleCase($Segment)
                }

                # Rejoin segments and replace the function name supplied by the user
                $Function = $CapitalizedName -join "-"
            }
            #endregion Input Validation: Name: Capitalization
            #endregion Input Validation: Name

            #region Produce Output String
            $Results = "Function $Function`n" + $FunctionBlock
            #endregion Produce Output String

            #region Output Handling
            #region Output Handling: OutFile
            IF ($PSCmdlet.ParameterSetName -eq "OutFile")
            {
                IF ($OutputPath)
                {
                    # Test whether output path is a file or directory
                    SWITCH -Regex ($OutputPath)
                    {
                        #region Output Handling: File
                        # Absolute regex madness
                        "\.\w\w$|\.\w\w\w$|\.\w\w\w\w$"
                        {
                            $OutputFilename  = $OutputPath.Split('\')[-1]
                            $OutputDirectory = $OutputPath.Substring(0,$OutputPath.LastIndexOf('\'))
                        }
                        #endregion Output Handling: File
                        #region Output Handling: Folder
                        Default
                        {
                            $OutputFilename  = "$Function.ps1"
                            $OutputDirectory = $OutputPath
                        }
                        #endregion Output Handling: Folder
                    }
                }
            }
            #endregion Output Handling: OutFile

            #region Output Handling: Module
            ELSEIF ($PSCmdlet.ParameterSetName -eq "Module")
            {
                TRY
                {
                    $OutputFilename  = "$Function.ps1"
                    $OutputDirectory = "$(Resolve-Path (Get-Module $Module -ListAvailable).ModuleBase)\$Visibility"
                }
                CATCH
                {
                    Write-Warning "Unable to determine function directory. Defaulting to present working directory: $PWD"
                    $OutputFilename  = "$Function.ps1"
                    $OutputDirectory = $PWD
                }
            }
            #endregion Output Handling: Module

            #region Output Handling: File Output
            IF ($PSCmdlet.ParameterSetName -match "OutFile|Module")
            {
                # Create path if it doesn't exist
                IF (!(Test-Path $OutputDirectory))
                {
                    New-Item -Path $OutputDirectory -ItemType Directory | Out-Null
                }

                # Output file
                $Results | Out-File $OutputDirectory\$OutputFilename -Encoding ascii -Force:$Force -Append:$Append
            }
            #endregion Output Handling: File Output
            #endregion Output Handling

            #region Results
            # Return results to the console if an output has not been specified or -Passthru has been selected
            IF(($PSCmdlet.ParameterSetName -notmatch "OutFile|Module") -or $Passthru)
            {
                Write-Output $Results
            }
            #endregion Results
        }
    }
    #endregion PROCESS Block

    #region END Block
    END
    {
        # Nothing to see here. Move along.
    }
    #endregion END Block
}