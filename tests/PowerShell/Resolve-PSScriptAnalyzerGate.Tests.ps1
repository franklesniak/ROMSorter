BeforeAll {
    $strRepositoryRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
    . (Join-Path -Path $strRepositoryRoot -ChildPath 'src/tools/Resolve-PSScriptAnalyzerGate.ps1')

    function Get-SyntheticAnalyzerFinding {
        # .SYNOPSIS
        # Creates a synthetic PSScriptAnalyzer finding for tests.
        #
        # .DESCRIPTION
        # Builds a PSCustomObject that contains the analyzer fields used by
        # Resolve-PSScriptAnalyzerGate. Tests use synthetic findings so they can
        # exercise Information and unknown severities without changing the
        # committed PSScriptAnalyzer settings.
        #
        # .PARAMETER Severity
        # The severity value to include in the finding.
        #
        # .PARAMETER RuleName
        # The rule name to include in the finding.
        #
        # .PARAMETER Message
        # The message to include in the finding.
        #
        # .PARAMETER ScriptPath
        # The script path to include in the finding.
        #
        # .PARAMETER Line
        # The line number to include in the finding.
        #
        # .PARAMETER Column
        # The column number to include in the finding.
        #
        # .EXAMPLE
        # Get-SyntheticAnalyzerFinding -Severity 'Warning'
        # # Returns a synthetic Warning finding.
        #
        # .INPUTS
        # None. This function does not accept pipeline input.
        #
        # .OUTPUTS
        # [pscustomobject] A synthetic analyzer finding.
        #
        # .NOTES
        # PRIVATE/INTERNAL HELPER - This function is not part of the public
        # API surface. Parameters, return shape, and positional contract may
        # change without notice.
        #
        # Version: 1.0.20260604.0
        # Positional parameters are not supported.
        #
        [CmdletBinding(PositionalBinding = $false)]
        [OutputType([pscustomobject])]
        param(
            [AllowNull()]
            [object]$Severity = 'Warning',

            [ValidateNotNullOrEmpty()]
            [string]$RuleName = 'PSExampleRule',

            [ValidateNotNullOrEmpty()]
            [string]$Message = 'Synthetic analyzer finding.',

            [ValidateNotNullOrEmpty()]
            [string]$ScriptPath = 'src/tools/TestScript.ps1',

            [int]$Line = 5,

            [int]$Column = 9
        )

        Set-StrictMode -Version Latest

        return [pscustomobject]@{
            Severity = $Severity
            RuleName = $RuleName
            Message = $Message
            ScriptPath = $ScriptPath
            Line = $Line
            Column = $Column
        }
    }
}

Describe "Resolve-PSScriptAnalyzerGate" {
    BeforeEach {
        $script:strOriginalGitHubActions = $env:GITHUB_ACTIONS
        $script:strOriginalTfBuild = $env:TF_BUILD
        $script:strOriginalBuildSourcesDirectory = $env:BUILD_SOURCESDIRECTORY
        $script:strOriginalGitHubWorkspace = $env:GITHUB_WORKSPACE
        Remove-Item -LiteralPath Env:\GITHUB_ACTIONS -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath Env:\TF_BUILD -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath Env:\BUILD_SOURCESDIRECTORY -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath Env:\GITHUB_WORKSPACE -ErrorAction SilentlyContinue
    }

    AfterEach {
        if ($null -eq $script:strOriginalGitHubActions) {
            Remove-Item -LiteralPath Env:\GITHUB_ACTIONS -ErrorAction SilentlyContinue
        } else {
            $env:GITHUB_ACTIONS = $script:strOriginalGitHubActions
        }

        if ($null -eq $script:strOriginalTfBuild) {
            Remove-Item -LiteralPath Env:\TF_BUILD -ErrorAction SilentlyContinue
        } else {
            $env:TF_BUILD = $script:strOriginalTfBuild
        }

        if ($null -eq $script:strOriginalBuildSourcesDirectory) {
            Remove-Item -LiteralPath Env:\BUILD_SOURCESDIRECTORY -ErrorAction SilentlyContinue
        } else {
            $env:BUILD_SOURCESDIRECTORY = $script:strOriginalBuildSourcesDirectory
        }

        if ($null -eq $script:strOriginalGitHubWorkspace) {
            Remove-Item -LiteralPath Env:\GITHUB_WORKSPACE -ErrorAction SilentlyContinue
        } else {
            $env:GITHUB_WORKSPACE = $script:strOriginalGitHubWorkspace
        }
    }

    Context "When resolving gate modes" {
        It "Resolves a missing mode to strict" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Information'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.Mode | Should -Be 'strict'
        }

        It "Resolves an empty mode to strict" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Information'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode '' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.Mode | Should -Be 'strict'
        }

        It "Resolves an unrecognized mode to strict" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Information'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'relaxed' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.Mode | Should -Be 'strict'
        }

        It "Resolves first-adoption mode when explicitly configured" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Information'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'first-adoption' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.Mode | Should -Be 'first-adoption'
        }
    }

    Context "When resolving annotation formats" {
        It "Falls back to plain output in Auto mode outside supported CI hosts" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Warning'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.AnnotationFormat | Should -Be 'Plain'
            $objResult.AnnotationCommands | Should -HaveCount 1
            $objResult.AnnotationCommands[0] | Should -Match '^PSScriptAnalyzer Warning:'
        }

        It "Detects GitHub Actions in Auto mode" {
            # Arrange
            $env:GITHUB_ACTIONS = 'true'
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Warning'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.AnnotationFormat | Should -Be 'GitHubActions'
            $objResult.AnnotationCommands | Should -HaveCount 1
            $objResult.AnnotationCommands[0] | Should -Match '^::warning '
        }

        It "Detects Azure Pipelines in Auto mode" {
            # Arrange
            $env:TF_BUILD = 'True'
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Warning'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.AnnotationFormat | Should -Be 'AzurePipelines'
            $objResult.AnnotationCommands | Should -HaveCount 1
            $objResult.AnnotationCommands[0] | Should -Match '^##vso\[task\.logissue type=warning;'
        }

        It "Lets an explicit GitHub Actions format override Azure Pipelines detection" {
            # Arrange
            $env:TF_BUILD = 'True'
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Warning'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate `
                -Mode 'strict' `
                -AnalyzerFinding @($objFinding) `
                -AnnotationFormat 'GitHubActions'

            # Assert
            $objResult.AnnotationFormat | Should -Be 'GitHubActions'
            $objResult.AnnotationCommands | Should -HaveCount 1
            $objResult.AnnotationCommands[0] | Should -Match '^::warning '
        }

        It "Lets an explicit Azure Pipelines format override GitHub Actions detection" {
            # Arrange
            $env:GITHUB_ACTIONS = 'true'
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Warning'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate `
                -Mode 'strict' `
                -AnalyzerFinding @($objFinding) `
                -AnnotationFormat 'AzurePipelines'

            # Assert
            $objResult.AnnotationFormat | Should -Be 'AzurePipelines'
            $objResult.AnnotationCommands | Should -HaveCount 1
            $objResult.AnnotationCommands[0] | Should -Match '^##vso\[task\.logissue type=warning;'
        }

        It "Fails closed when Auto mode sees conflicting CI host indicators" {
            # Arrange
            $env:GITHUB_ACTIONS = 'true'
            $env:TF_BUILD = 'True'
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Warning'

            # Act / Assert
            {
                Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding @($objFinding)
            } | Should -Throw -ExpectedMessage '*Pass an explicit -AnnotationFormat value*'
        }
    }

    Context "When strict mode gates severities" {
        It "Fails Error findings" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Error'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.ShouldFail | Should -BeTrue
            $objResult.Summary.BlockingCount | Should -Be 1
        }

        It "Fails Warning findings" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Warning'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.ShouldFail | Should -BeTrue
            $objResult.Summary.BlockingCount | Should -Be 1
        }

        It "Allows Information findings" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Information'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.ShouldFail | Should -BeFalse
            $objResult.Summary.BlockingCount | Should -Be 0
        }

        It "Fails unknown-severity findings" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Audit'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.ShouldFail | Should -BeTrue
            $objResult.Summary.UnknownSeverityCount | Should -Be 1
        }
    }

    Context "When first-adoption mode gates severities" {
        It "Fails Error findings" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Error'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'first-adoption' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.ShouldFail | Should -BeTrue
            $objResult.Summary.BlockingCount | Should -Be 1
        }

        It "Allows Warning findings as tracked debt" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Warning'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'first-adoption' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.ShouldFail | Should -BeFalse
            $objResult.Summary.BlockingCount | Should -Be 0
            $objResult.Summary.DebtCount | Should -Be 1
            $objResult.Findings.Count | Should -Be 1
            $objResult.Findings[0].TrackedDebt | Should -BeTrue
        }

        It "Allows Information findings as tracked debt" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Information'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'first-adoption' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.ShouldFail | Should -BeFalse
            $objResult.Summary.BlockingCount | Should -Be 0
            $objResult.Summary.DebtCount | Should -Be 1
            $objResult.Findings.Count | Should -Be 1
            $objResult.Findings[0].TrackedDebt | Should -BeTrue
        }

        It "Fails unknown-severity findings" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Audit'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'first-adoption' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.ShouldFail | Should -BeTrue
            $objResult.Summary.UnknownSeverityCount | Should -Be 1
        }
    }

    Context "When findings cannot be parsed normally" {
        It "Treats a missing severity as unknown in strict mode" {
            # Arrange
            $objFinding = [pscustomobject]@{
                RuleName = 'PSMissingSeverity'
                Message = 'Missing severity.'
                ScriptPath = 'src/tools/TestScript.ps1'
                Line = 3
                Column = 7
            }

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.ShouldFail | Should -BeTrue
            $objResult.Findings.Count | Should -Be 1
            $objResult.Findings[0].Severity | Should -Be 'Unknown'
        }

        It "Treats a missing severity as unknown in first-adoption mode" {
            # Arrange
            $objFinding = [pscustomobject]@{
                RuleName = 'PSMissingSeverity'
                Message = 'Missing severity.'
                ScriptPath = 'src/tools/TestScript.ps1'
                Line = 3
                Column = 7
            }

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'first-adoption' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.ShouldFail | Should -BeTrue
            $objResult.Findings.Count | Should -Be 1
            $objResult.Findings[0].Severity | Should -Be 'Unknown'
        }
    }

    Context "When building annotations and summaries" {
        It "Creates one GitHub Actions annotation command for every finding" {
            # Arrange
            $arrFinding = @(
                Get-SyntheticAnalyzerFinding -Severity 'Error'
                Get-SyntheticAnalyzerFinding -Severity 'Warning'
                Get-SyntheticAnalyzerFinding -Severity 'Information'
                Get-SyntheticAnalyzerFinding -Severity 'Audit'
            )

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate `
                -Mode 'first-adoption' `
                -AnalyzerFinding $arrFinding `
                -AnnotationFormat 'GitHubActions'

            # Assert
            $objResult.AnnotationFormat | Should -Be 'GitHubActions'
            $objResult.AnnotationCommands | Should -HaveCount 4
            $objResult.AnnotationCommands[0] | Should -Match '^::error '
            $objResult.AnnotationCommands[1] | Should -Match '^::warning '
            $objResult.AnnotationCommands[2] | Should -Match '^::notice '
            $objResult.AnnotationCommands[3] | Should -Match '^::error '
        }

        It "Prints a first-adoption warning debt summary with Information counts" {
            # Arrange
            $arrFinding = @(
                Get-SyntheticAnalyzerFinding -Severity 'Warning'
                Get-SyntheticAnalyzerFinding -Severity 'Warning'
                Get-SyntheticAnalyzerFinding -Severity 'Information'
            )

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'first-adoption' -AnalyzerFinding $arrFinding

            # Assert
            $objResult.ShouldFail | Should -BeFalse
            $objResult.Summary.DebtCount | Should -Be 3
            ($objResult.SummaryLines -join "`n") | Should -Match 'Warning debt: 2 Warning and 1 Information'
            ($objResult.SummaryLines -join "`n") | Should -Match '_TODO-repo-init\.md or a post-adoption issue'
        }

        It "Prints a strict summary that identifies blocking findings" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Warning'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.ShouldFail | Should -BeTrue
            ($objResult.SummaryLines -join "`n") | Should -Match 'Strict gate'
            ($objResult.SummaryLines -join "`n") | Should -Match 'Result: fail'
        }

        It "Returns top rule and file summaries in descending count order" {
            # Arrange
            $arrFinding = @(
                Get-SyntheticAnalyzerFinding -Severity 'Warning' -RuleName 'PSRuleB' -ScriptPath 'src/tools/B.ps1'
                Get-SyntheticAnalyzerFinding -Severity 'Warning' -RuleName 'PSRuleA' -ScriptPath 'src/tools/B.ps1'
                Get-SyntheticAnalyzerFinding -Severity 'Information' -RuleName 'PSRuleB' -ScriptPath 'src/tools/A.ps1'
            )

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'first-adoption' -AnalyzerFinding $arrFinding

            # Assert
            $objResult.TopRules | Should -Not -BeNullOrEmpty
            $objResult.TopFiles | Should -Not -BeNullOrEmpty
            $objResult.TopRules[0].RuleName | Should -Be 'PSRuleB'
            $objResult.TopRules[0].Count | Should -Be 2
            $objResult.TopRules[1].RuleName | Should -Be 'PSRuleA'
            $objResult.TopFiles[0].ScriptPath | Should -Be 'src/tools/B.ps1'
            $objResult.TopFiles[0].Count | Should -Be 2
            ($objResult.SummaryLines -join "`n") | Should -Match 'Top rule findings: PSRuleB \(2\); PSRuleA \(1\)'
            ($objResult.SummaryLines -join "`n") | Should -Match 'Top files by findings: src/tools/B\.ps1 \(2\); src/tools/A\.ps1 \(1\)'
        }

        It "Recommends first-adoption mode and emits issue-ready Markdown for warning debt" {
            # Arrange
            $arrFinding = @(
                Get-SyntheticAnalyzerFinding -Severity 'Warning' -RuleName 'PSRuleA'
                Get-SyntheticAnalyzerFinding -Severity 'Information' -RuleName 'PSRuleB'
            )

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding $arrFinding

            # Assert
            $objResult.RecommendedMode | Should -Be 'first-adoption'
            $objResult.Summary.RecommendedMode | Should -Be 'first-adoption'
            $objResult.IssueReadyMarkdown | Should -Not -BeNullOrEmpty
            ($objResult.IssueReadyMarkdown -join "`n") | Should -Match 'PSScriptAnalyzer First-Adoption Debt Cleanup'
            ($objResult.IssueReadyMarkdown -join "`n") | Should -Match 'PSSCRIPTANALYZER_GATE_MODE'
        }

        It "Recommends strict mode when errors or unknown severities are present" {
            # Arrange
            $arrFinding = @(
                Get-SyntheticAnalyzerFinding -Severity 'Error'
                Get-SyntheticAnalyzerFinding -Severity 'Audit'
            )

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'first-adoption' -AnalyzerFinding $arrFinding

            # Assert
            $objResult.RecommendedMode | Should -Be 'strict'
            $objResult.IssueReadyMarkdown | Should -BeNullOrEmpty
        }
    }

    Context "When rendering Azure Pipelines and plain output" {
        It "Escapes Azure property values separately from message values" {
            # Arrange
            $strValue = "a;b]c%d`r`ne"

            # Act
            $strPropertyValue = ConvertTo-AzurePipelinesLoggingCommandPropertyValue -Value $strValue
            $strMessageValue = ConvertTo-AzurePipelinesLoggingCommandMessage -Value $strValue

            # Assert
            $strPropertyValue | Should -Be 'a%3Bb%5Dc%AZP25d%0D%0Ae'
            $strMessageValue | Should -Be 'a;b]c%AZP25d%0D%0Ae'
        }

        It "Maps Azure severities without promoting Information findings" {
            # Arrange
            $arrFinding = @(
                Get-SyntheticAnalyzerFinding -Severity 'Error' -RuleName 'PSErrorRule'
                Get-SyntheticAnalyzerFinding -Severity 'Warning' -RuleName 'PSWarningRule'
                Get-SyntheticAnalyzerFinding -Severity 'Information' -RuleName 'PSInformationRule'
                Get-SyntheticAnalyzerFinding -Severity 'Audit' -RuleName 'PSUnknownRule'
            )

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate `
                -Mode 'first-adoption' `
                -AnalyzerFinding $arrFinding `
                -AnnotationFormat 'AzurePipelines'

            # Assert
            $objResult.AnnotationFormat | Should -Be 'AzurePipelines'
            $objResult.AnnotationCommands | Should -HaveCount 4
            $objResult.AnnotationCommands[0] | Should -Match '^##vso\[task\.logissue type=error;'
            $objResult.AnnotationCommands[0] | Should -Match 'code=PSErrorRule;'
            $objResult.AnnotationCommands[1] | Should -Match '^##vso\[task\.logissue type=warning;'
            $objResult.AnnotationCommands[1] | Should -Match 'code=PSWarningRule;'
            $objResult.AnnotationCommands[2] | Should -Match '^PSScriptAnalyzer Information:'
            $objResult.AnnotationCommands[2] | Should -Not -Match '^##vso'
            $objResult.AnnotationCommands[3] | Should -Match '^##vso\[task\.logissue type=error;'
            $objResult.AnnotationCommands[3] | Should -Match '\[Unknown \(Audit\)\]'
        }

        It "Escapes Azure logging-command properties without over-escaping messages" {
            # Arrange
            # The property/message escaping split mirrors the Azure Pipelines
            # task library implementation:
            # https://github.com/microsoft/azure-pipelines-task-lib/blob/990d322531e994631423bf01603a5ae10b59d96c/node/taskcommand.ts#L622-L651
            $objFinding = Get-SyntheticAnalyzerFinding `
                -Severity 'Warning' `
                -RuleName 'PSRule;One]100%' `
                -Message "Use ; and ] but 100% only.`nNext line." `
                -ScriptPath '/agent/_work/1/s/src/tools/Has;Bracket].ps1'

            $strExpected = (
                '##vso[task.logissue type=warning;sourcepath=/agent/_work/1/s/src/tools/Has%3BBracket%5D.ps1;' +
                'linenumber=5;columnnumber=9;code=PSRule%3BOne%5D100%AZP25;]' +
                '[Warning] PSRule;One]100%AZP25 - Use ; and ] but 100%AZP25 only.%0ANext line.'
            )

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate `
                -Mode 'strict' `
                -AnalyzerFinding @($objFinding) `
                -AnnotationFormat 'AzurePipelines'

            # Assert
            $objResult.AnnotationCommands | Should -HaveCount 1
            $objResult.AnnotationCommands[0] | Should -Be $strExpected
            $objResult.AnnotationCommands[0].Contains("`r") | Should -BeFalse
            $objResult.AnnotationCommands[0].Contains("`n") | Should -BeFalse

            $intMessageStartIndex = $objResult.AnnotationCommands[0].IndexOf(']') + 1
            $strRenderedMessage = $objResult.AnnotationCommands[0].Substring($intMessageStartIndex)
            $strRenderedMessage | Should -Not -Match '%3B|%5D'
        }

        It "Uses an absolute Azure source path when a repository root is available" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding `
                -Severity 'Warning' `
                -ScriptPath 'src/tools/Demo.ps1'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate `
                -Mode 'strict' `
                -RepositoryRoot '/agent/_work/1/s' `
                -AnalyzerFinding @($objFinding) `
                -AnnotationFormat 'AzurePipelines'

            # Assert
            $objResult.Findings.Count | Should -Be 1
            $objResult.Findings[0].ScriptPath | Should -Be 'src/tools/Demo.ps1'
            $objResult.Findings[0].AzureSourcePath | Should -Be '/agent/_work/1/s/src/tools/Demo.ps1'
            $objResult.AnnotationCommands | Should -HaveCount 1
            $objResult.AnnotationCommands[0] | Should -Match 'sourcepath=/agent/_work/1/s/src/tools/Demo\.ps1;'
        }

        It "Uses BUILD_SOURCESDIRECTORY as the Azure repository root when no root is passed" {
            # Arrange
            $env:TF_BUILD = 'True'
            $env:BUILD_SOURCESDIRECTORY = '/agent/_work/1/s'
            $objFinding = Get-SyntheticAnalyzerFinding `
                -Severity 'Warning' `
                -ScriptPath 'src/tools/Demo.ps1'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.AnnotationFormat | Should -Be 'AzurePipelines'
            $objResult.Findings.Count | Should -Be 1
            $objResult.Findings[0].AzureSourcePath | Should -Be '/agent/_work/1/s/src/tools/Demo.ps1'
            $objResult.AnnotationCommands[0] | Should -Match 'sourcepath=/agent/_work/1/s/src/tools/Demo\.ps1;'
        }

        It "Renders plain output as a single line per finding" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding `
                -Severity 'Warning' `
                -Message "Line one.`nLine two."

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate `
                -Mode 'strict' `
                -AnalyzerFinding @($objFinding) `
                -AnnotationFormat 'Plain'

            # Assert
            $objResult.AnnotationFormat | Should -Be 'Plain'
            $objResult.AnnotationCommands | Should -HaveCount 1
            $objResult.AnnotationCommands[0] | Should -Be 'PSScriptAnalyzer Warning: src/tools/TestScript.ps1:5:9 [PSExampleRule] Line one. Line two.'
            $objResult.AnnotationCommands[0] | Should -Not -Match '^::|^##vso'
            $objResult.AnnotationCommands[0].Contains("`r") | Should -BeFalse
            $objResult.AnnotationCommands[0].Contains("`n") | Should -BeFalse
        }
    }

    Context "When normalizing annotation paths" {
        It "Makes an absolute ScriptPath relative to the repository root" {
            # Arrange
            $strRepositoryRoot = '/home/runner/work/repo/repo'
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Warning' -ScriptPath "$strRepositoryRoot/src/tools/Demo.ps1"

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate `
                -Mode 'strict' `
                -RepositoryRoot $strRepositoryRoot `
                -AnalyzerFinding @($objFinding) `
                -AnnotationFormat 'GitHubActions'

            # Assert
            $objResult.Findings.Count | Should -Be 1
            $objResult.Findings[0].ScriptPath | Should -Be 'src/tools/Demo.ps1'
            ($objResult.AnnotationCommands -join "`n") | Should -Match 'file=src/tools/Demo\.ps1'
        }

        It "Leaves a ScriptPath unchanged when it is outside the repository root" {
            # Arrange
            $objFinding = Get-SyntheticAnalyzerFinding -Severity 'Warning' -ScriptPath '/var/other/Demo.ps1'

            # Act
            $objResult = Resolve-PSScriptAnalyzerGate -Mode 'strict' -RepositoryRoot '/home/runner/work/repo/repo' -AnalyzerFinding @($objFinding)

            # Assert
            $objResult.Findings.Count | Should -Be 1
            $objResult.Findings[0].ScriptPath | Should -Be '/var/other/Demo.ps1'
        }
    }
}
