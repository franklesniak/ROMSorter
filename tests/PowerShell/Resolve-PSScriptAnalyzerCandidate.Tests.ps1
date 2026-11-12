BeforeAll {
    $strRepositoryRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
    . (Join-Path -Path $strRepositoryRoot -ChildPath 'src/tools/Resolve-PSScriptAnalyzerCandidate.ps1')

    function Initialize-CandidateTestRepository {
        # .SYNOPSIS
        # Creates a temporary repository-like directory for candidate tests.
        #
        # .DESCRIPTION
        # Creates a unique directory under the system temporary directory. Tests
        # use this helper so filesystem fixtures are isolated and easy to clean
        # up after each test case.
        #
        # .EXAMPLE
        # $strRoot = Initialize-CandidateTestRepository
        #
        # # Returns a unique temporary directory path.
        #
        # .INPUTS
        # None. This function does not accept pipeline input.
        #
        # .OUTPUTS
        # [string] The temporary repository root path.
        #
        # .NOTES
        # PRIVATE/INTERNAL HELPER - This function is not part of the public
        # API surface. Parameters, return shape, and positional contract may
        # change without notice.
        #
        # Version: 1.0.20260701.0
        # Positional parameters are not supported.
        #
        [CmdletBinding(PositionalBinding = $false)]
        [OutputType([string])]
        param()

        Set-StrictMode -Version Latest

        $strRoot = [System.IO.Path]::Combine(
            [System.IO.Path]::GetTempPath(),
            ('pssa-candidate-{0}' -f [System.Guid]::NewGuid().ToString('N'))
        )
        [void]([System.IO.Directory]::CreateDirectory($strRoot))
        return $strRoot
    }

    function Add-CandidateTestFile {
        # .SYNOPSIS
        # Adds a UTF-8 test file beneath a temporary repository root.
        #
        # .DESCRIPTION
        # Creates parent directories as needed and writes a small PowerShell
        # source file using an explicit UTF-8 encoding.
        #
        # .PARAMETER RepositoryRoot
        # The temporary repository root.
        #
        # .PARAMETER RepositoryRelativePath
        # The repository-relative file path to create.
        #
        # .EXAMPLE
        # Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath 'src/a.ps1'
        #
        # # Creates src/a.ps1 under the repository root.
        #
        # .INPUTS
        # None. This function does not accept pipeline input.
        #
        # .OUTPUTS
        # None.
        #
        # .NOTES
        # PRIVATE/INTERNAL HELPER - This function is not part of the public
        # API surface. Parameters, return shape, and positional contract may
        # change without notice.
        #
        # Version: 1.0.20260701.0
        # Positional parameters are not supported.
        #
        [CmdletBinding(PositionalBinding = $false)]
        [OutputType([void])]
        param(
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$RepositoryRoot,

            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [string]$RepositoryRelativePath
        )

        Set-StrictMode -Version Latest

        $strFullPath = [System.IO.Path]::Combine($RepositoryRoot, $RepositoryRelativePath)
        $strParentDirectory = [System.IO.Path]::GetDirectoryName($strFullPath)
        if (-not [string]::IsNullOrEmpty($strParentDirectory)) {
            [void]([System.IO.Directory]::CreateDirectory($strParentDirectory))
        }
        $objEncoding = [System.Text.UTF8Encoding]::new($false)
        [System.IO.File]::WriteAllText($strFullPath, "Write-Output 'test'`n", $objEncoding)
    }
}

Describe "Resolve-PSScriptAnalyzerCandidate" {
    It "Selects PS1, PSM1, and PSD1 candidates case-insensitively" {
        # Arrange
        $strRoot = Initialize-CandidateTestRepository
        try {
            $arrAllowedExtension = @(Get-PSScriptAnalyzerCandidateAllowedExtension)
            $arrRelativePath = @(
                'src/Script.PS1'
                'src/Module.PSM1'
                'src/Data.PSD1'
            )
            foreach ($strRelativePath in $arrRelativePath) {
                Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath $strRelativePath
            }

            # Act
            $arrCandidate = @(
                foreach ($strRelativePath in $arrRelativePath) {
                    Resolve-PSScriptAnalyzerCandidate `
                        -RepositoryRoot $strRoot `
                        -CandidatePath $strRelativePath `
                        -RepositoryRelativePath $strRelativePath
                }
            )

            # Assert
            $arrAllowedExtension | Should -Contain '.ps1'
            $arrAllowedExtension | Should -Contain '.psm1'
            $arrAllowedExtension | Should -Contain '.psd1'
            $arrAllowedExtension | Should -HaveCount 3
            $arrCandidate | Should -HaveCount 3
            foreach ($objCandidate in $arrCandidate) {
                $objCandidate.OutcomeCategory | Should -Be 'selected'
                $objCandidate.ReasonCode | Should -Be 'Selected'
            }
        } finally {
            Remove-Item -LiteralPath $strRoot -Recurse -Force
        }
    }

    It "Selects bracket-containing paths and escapes analyzer input literally" {
        # Arrange
        $strRoot = Initialize-CandidateTestRepository
        try {
            $strRelativePath = 'src/Invoke-[Build].ps1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath $strRelativePath
            $strFullPath = Resolve-PSScriptAnalyzerFullPath -RepositoryRoot $strRoot -Path $strRelativePath

            # Act
            $objCandidate = Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRoot `
                -CandidatePath $strRelativePath `
                -RepositoryRelativePath $strRelativePath

            # Assert
            $objCandidate.OutcomeCategory | Should -Be 'selected'
            $objCandidate.RepositoryRelativePath | Should -Be $strRelativePath
            $objCandidate.EscapedAnalyzerPath | Should -Be (
                [System.Management.Automation.WildcardPattern]::Escape($strFullPath)
            )
        } finally {
            Remove-Item -LiteralPath $strRoot -Recurse -Force
        }
    }

    It "Excludes exact node_modules segments without excluding substrings" {
        # Arrange
        $strRoot = Initialize-CandidateTestRepository
        try {
            $strExcludedPath = 'node_modules/.bin/tool.ps1'
            $strSelectedPath = 'tools/node_modules_helper/Build.ps1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath $strExcludedPath
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath $strSelectedPath

            # Act
            $objExcludedCandidate = Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRoot `
                -CandidatePath $strExcludedPath `
                -RepositoryRelativePath $strExcludedPath
            $objSelectedCandidate = Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRoot `
                -CandidatePath $strSelectedPath `
                -RepositoryRelativePath $strSelectedPath

            # Assert
            $objExcludedCandidate.OutcomeCategory | Should -Be 'policy-excluded'
            $objExcludedCandidate.ReasonCode | Should -Be 'NodeModulesSegment'
            $objSelectedCandidate.OutcomeCategory | Should -Be 'selected'
        } finally {
            Remove-Item -LiteralPath $strRoot -Recurse -Force
        }
    }

    It "Excludes exact .git segments without excluding substrings" {
        # Arrange
        $strRoot = Initialize-CandidateTestRepository
        try {
            $strExcludedPath = '.git/hooks/pre-commit.ps1'
            $strSelectedPath = 'tools/git-helper/pre-commit.ps1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath $strExcludedPath
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath $strSelectedPath

            # Act
            $objExcludedCandidate = Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRoot `
                -CandidatePath $strExcludedPath `
                -RepositoryRelativePath $strExcludedPath
            $objSelectedCandidate = Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRoot `
                -CandidatePath $strSelectedPath `
                -RepositoryRelativePath $strSelectedPath

            # Assert
            $objExcludedCandidate.OutcomeCategory | Should -Be 'policy-excluded'
            $objExcludedCandidate.ReasonCode | Should -Be 'GitDirectorySegment'
            $objSelectedCandidate.OutcomeCategory | Should -Be 'selected'
        } finally {
            Remove-Item -LiteralPath $strRoot -Recurse -Force
        }
    }

    It "Excludes only the repository analyzer settings file" {
        # Arrange
        $strRoot = Initialize-CandidateTestRepository
        try {
            $strSettingsPath = '.github/linting/PSScriptAnalyzerSettings.psd1'
            $strSameBasenamePath = 'tools/PSScriptAnalyzerSettings.psd1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath $strSettingsPath
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath $strSameBasenamePath

            # Act
            $objSettingsCandidate = Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRoot `
                -CandidatePath $strSettingsPath `
                -RepositoryRelativePath $strSettingsPath
            $objSameBasenameCandidate = Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRoot `
                -CandidatePath $strSameBasenamePath `
                -RepositoryRelativePath $strSameBasenamePath

            # Assert
            $objSettingsCandidate.OutcomeCategory | Should -Be 'policy-excluded'
            $objSettingsCandidate.ReasonCode | Should -Be 'AnalyzerSettingsFile'
            $objSameBasenameCandidate.OutcomeCategory | Should -Be 'selected'
        } finally {
            Remove-Item -LiteralPath $strRoot -Recurse -Force
        }
    }

    It "Uses force-visible discovery while pruning .git and node_modules before descent" {
        # Arrange
        $strRoot = Initialize-CandidateTestRepository
        try {
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath '.hidden/Visible.psm1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath 'data/Manifest.psd1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath 'src/Build.ps1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath 'tools/PSScriptAnalyzerSettings.psd1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath '.github/linting/PSScriptAnalyzerSettings.psd1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath '.git/objects/aa/ignored.ps1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath 'node_modules/.bin/tool.ps1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath 'tools/node_modules_helper/Keep.ps1'

            # Act
            $arrCandidate = @(
                Get-PSScriptAnalyzerCandidate -RepositoryRoot $strRoot -DirectoryVisibility All
            )
            $arrPath = @($arrCandidate | ForEach-Object { $_.RepositoryRelativePath })
            $objSettingsCandidate = @(
                $arrCandidate |
                    Where-Object { $_.RepositoryRelativePath -eq '.github/linting/PSScriptAnalyzerSettings.psd1' }
            )

            # Assert
            $arrPath | Should -Contain '.hidden/Visible.psm1'
            $arrPath | Should -Contain 'data/Manifest.psd1'
            $arrPath | Should -Contain 'src/Build.ps1'
            $arrPath | Should -Contain 'tools/PSScriptAnalyzerSettings.psd1'
            $arrPath | Should -Contain 'tools/node_modules_helper/Keep.ps1'
            $arrPath | Should -Contain '.github/linting/PSScriptAnalyzerSettings.psd1'
            $arrPath | Should -Not -Contain '.git/objects/aa/ignored.ps1'
            $arrPath | Should -Not -Contain 'node_modules/.bin/tool.ps1'
            $objSettingsCandidate | Should -HaveCount 1
            $objSettingsCandidate[0].OutcomeCategory | Should -Be 'policy-excluded'
            $objSettingsCandidate[0].ReasonCode | Should -Be 'AnalyzerSettingsFile'
        } finally {
            Remove-Item -LiteralPath $strRoot -Recurse -Force
        }
    }

    It "Preserves CRLF in structured records while rendering summaries on one line" {
        # Arrange
        $strRelativePath = "src/line`r`nbreak.ps1"
        $objCandidate = ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName "/repo/$strRelativePath" `
            -RepositoryRelativePath $strRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode 'MissingTarget'

        # Act
        $objSummary = Get-PSScriptAnalyzerCandidateSummary -Candidate @($objCandidate)
        $arrSummaryLine = @(ConvertTo-PSScriptAnalyzerCandidateSummaryLine -CandidateSummary $objSummary)

        # Assert
        $objSummary.Unsafe | Should -Not -BeNullOrEmpty
        $objSummary.Unsafe[0].RepositoryRelativePath | Should -Be $strRelativePath
        ($arrSummaryLine -join "`n") | Should -Match 'src/line\\r\\nbreak\.ps1'
        foreach ($strLine in $arrSummaryLine) {
            $strLine.Contains("`r") | Should -BeFalse
            $strLine.Contains("`n") | Should -BeFalse
        }
    }

    It "Fails closed when a candidate has an unexpected OutcomeCategory" {
        # Arrange
        $objCandidate = [pscustomobject]@{ OutcomeCategory = 'unexpected-category' }

        # Act / Assert
        { Get-PSScriptAnalyzerCandidateSummary -Candidate @($objCandidate) } |
            Should -Throw -ExpectedMessage '*Unexpected PSScriptAnalyzer candidate OutcomeCategory*'
    }

    It "Classifies leaf reparse candidates that escape the repository as unsafe" {
        # Arrange
        $strRoot = Initialize-CandidateTestRepository
        $strExternalRoot = Initialize-CandidateTestRepository
        try {
            Add-CandidateTestFile -RepositoryRoot $strExternalRoot -RepositoryRelativePath 'outside.ps1'
            $strLinkPath = [System.IO.Path]::Combine($strRoot, 'link.ps1')
            $strTargetPath = [System.IO.Path]::Combine($strExternalRoot, 'outside.ps1')

            try {
                [void](New-Item -Path $strLinkPath -ItemType SymbolicLink -Target $strTargetPath -ErrorAction Stop)
            } catch {
                Set-ItResult -Skipped -Because 'Filesystem or current privileges do not allow symlink creation.'
                return
            }

            # Act
            $objCandidate = Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRoot `
                -CandidatePath 'link.ps1' `
                -RepositoryRelativePath 'link.ps1'

            # Assert
            $objCandidate.OutcomeCategory | Should -Be 'unsafe'
            $objCandidate.ReasonCode | Should -BeIn @(
                'TargetOutsideRepository',
                'ReparsePointResolutionUnsupported'
            )
        } finally {
            Remove-Item -LiteralPath $strRoot -Recurse -Force
            Remove-Item -LiteralPath $strExternalRoot -Recurse -Force
        }
    }

    It "Classifies leaf reparse candidates that resolve to policy-excluded targets as unsafe" {
        # Arrange
        $strRoot = Initialize-CandidateTestRepository
        try {
            $strSettingsPath = '.github/linting/PSScriptAnalyzerSettings.psd1'
            $strNodeModulesTargetPath = 'node_modules/.bin/target.ps1'
            $strGitTargetPath = '.git/hooks/target.ps1'
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath $strSettingsPath
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath $strNodeModulesTargetPath
            Add-CandidateTestFile -RepositoryRoot $strRoot -RepositoryRelativePath $strGitTargetPath

            $hashtableLinkTarget = @{
                'settings-link.ps1' = [System.IO.Path]::Combine($strRoot, $strSettingsPath)
                'node-link.ps1' = [System.IO.Path]::Combine($strRoot, $strNodeModulesTargetPath)
                'git-link.ps1' = [System.IO.Path]::Combine($strRoot, $strGitTargetPath)
            }

            try {
                foreach ($strLinkPath in $hashtableLinkTarget.Keys) {
                    $strFullLinkPath = [System.IO.Path]::Combine($strRoot, $strLinkPath)
                    [void](
                        New-Item `
                            -Path $strFullLinkPath `
                            -ItemType SymbolicLink `
                            -Target $hashtableLinkTarget[$strLinkPath] `
                            -ErrorAction Stop
                    )
                }
            } catch {
                Set-ItResult -Skipped -Because 'Filesystem or current privileges do not allow symlink creation.'
                return
            }

            # Act
            $objSettingsCandidate = Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRoot `
                -CandidatePath 'settings-link.ps1' `
                -RepositoryRelativePath 'settings-link.ps1'
            $objNodeModulesCandidate = Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRoot `
                -CandidatePath 'node-link.ps1' `
                -RepositoryRelativePath 'node-link.ps1'
            $objGitCandidate = Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRoot `
                -CandidatePath 'git-link.ps1' `
                -RepositoryRelativePath 'git-link.ps1'

            # Assert
            $objSettingsCandidate.OutcomeCategory | Should -Be 'unsafe'
            $objSettingsCandidate.ReasonCode | Should -BeIn @(
                'TargetAnalyzerSettingsFile',
                'ReparsePointResolutionUnsupported'
            )
            $objNodeModulesCandidate.OutcomeCategory | Should -Be 'unsafe'
            $objNodeModulesCandidate.ReasonCode | Should -BeIn @(
                'TargetNodeModulesSegment',
                'ReparsePointResolutionUnsupported'
            )
            $objGitCandidate.OutcomeCategory | Should -Be 'unsafe'
            $objGitCandidate.ReasonCode | Should -BeIn @(
                'TargetGitDirectorySegment',
                'ReparsePointResolutionUnsupported'
            )
        } finally {
            Remove-Item -LiteralPath $strRoot -Recurse -Force
        }
    }
}
