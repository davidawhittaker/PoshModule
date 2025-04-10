# Get public and private function files
$Public  = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1  -ErrorAction SilentlyContinue)
$Private = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue)

$Manifest = Get-ChildItem -Path $PSScriptRoot\*.psd1

# Dot source the files
FOREACH ($Import in ($Public + $Private))
{
    TRY
    {
        . $Import.FullName
    }
    CATCH
    {
        Write-Error -Message "Failed to import function $($Import.FullName): $_"
    }
}

# Update module manifest on each subsequent reload with any new .ps1 files that have been created in the public directory
Update-ModuleManifest -FunctionsToExport $Public.BaseName -Path $Manifest.FullName