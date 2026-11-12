# Parse each public script without executing downloads or file mutations.
# Version: 1.0.20261111.0
# Last Updated: 2026-11-11

Describe 'Public script syntax on the current PowerShell host' {
    It 'Parses <Name> without executing it' -ForEach @(
        Get-ChildItem -LiteralPath (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) -Filter '*.ps1' -File |
            ForEach-Object { @{ Name = $_.Name; FullName = $_.FullName } }
    ) {
        $arrTokens = $null
        $arrParseErrors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile(
            $FullName, [ref]$arrTokens, [ref]$arrParseErrors
        )
        @($arrParseErrors).Count | Should -Be 0 -Because ($arrParseErrors -join '; ')
    }
}
