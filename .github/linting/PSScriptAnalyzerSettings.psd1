# PSScriptAnalyzer Settings for Repository Template
# Enforces OTBS (One True Brace Style) / K&R formatting
# Aligns with .github/instructions/powershell.instructions.md
#
# Version: 1.0.20260430.0
#
# Usage:
#   Invoke-ScriptAnalyzer -Path .\script.ps1 -Settings .\.github\linting\PSScriptAnalyzerSettings.psd1
#   Invoke-ScriptAnalyzer -Path .\script.ps1 -Settings .\.github\linting\PSScriptAnalyzerSettings.psd1 -Fix
#
# For more information on PSScriptAnalyzer rules, see:
#   https://github.com/PowerShell/PSScriptAnalyzer

@{
    # Include default rules
    IncludeDefaultRules = $true

    # Severity levels to enforce
    Severity = @('Error', 'Warning')

    # Code formatting rules (OTBS / K&R style)
    Rules = @{
        PSPlaceOpenBrace = @{
            # Turn the rule on
            Enable = $true
            OnSameLine = $true
            NewLineAfter = $true

            # Allow small exceptions, e.g.:  if ($true) { "blah" } else { "blah blah" }
            IgnoreOneLineBlock = $true
        }
        # Closing braces must not have empty lines before them per OTBS style.
        # This keeps code compact and aligns with community best practices.
        PSPlaceCloseBrace = @{
            # Turn the rule on
            Enable = $true

            # Minimize whitespace
            NoEmptyLineBefore = $true

            # Allow small exceptions, e.g.:  if ($true) { "blah" } else { "blah blah" }
            IgnoreOneLineBlock = $true

            NewLineAfter = $false
        }
        PSUseConsistentIndentation = @{
            # Turn the rule on
            Enable = $true

            # Each indentation should be 4 spaces
            IndentationSize = 4
            Kind = 'space'

            # Multi-line pipeline statements should indent every line following the second only once
            PipelineIndentation = 'IncreaseIndentationForFirstPipeline'
        }
        PSUseConsistentWhitespace = @{
            # Turn the rule on
            Enable = $true

            # e.g.: enforce: if ($true) { foo } instead of:  if ($true) {bar}
            CheckInnerBrace = $true

            # e.g.: enforce: foo { } instead of: foo{ }
            CheckOpenBrace = $true

            # e.g.: enforce: if (true) instead of: if(true)
            CheckOpenParen = $true

            # e.g.: enforce: $x = 1 instead of: $x=1
            CheckOperator = $true

            # e.g.: enforce: foo | bar instead of: foo|bar
            CheckPipe = $true

            # e.g.: enforce: @(1, 2, 3) or @{a = 1; b = 2} instead of: @(1,2,3) or @{a = 1;b = 2}
            CheckSeparator = $true

            # Do not flag multiple spaces before a pipe as redundant
            CheckPipeForRedundantWhitespace = $false

            # Do not enforce whitespace around parameter names
            CheckParameter = $false

            # Enforce whitespace around assignment operators inside hashtables
            IgnoreAssignmentOperatorInsideHashTable = $false
        }

        # Disable vertical alignment (per style guide)
        PSAlignAssignmentStatement = @{
            Enable = $false
            CheckHashtable = $false
        }

        # Avoid using aliases
        PSAvoidUsingCmdletAliases = @{
            Enable = $true
        }

        # Use approved verbs
        PSUseApprovedVerbs = @{
            Enable = $true
        }

        # Use singular nouns
        PSUseSingularNouns = @{
            Enable = $true
        }

        # Enforce literal hashtable initialization (use @{} instead of New-Object)
        PSUseLiteralInitializerForHashtable = @{
            Enable = $true
        }

        # Provide comment-based help
        PSProvideCommentHelp = @{
            Enable = $true
            ExportedOnly = $false
            BlockComment = $false
            VSCodeSnippetCorrection = $false
            Placement = 'begin'
        }

        # Encourage named parameters in function calls for readability
        PSAvoidUsingPositionalParameters = @{
            Enable = $true
        }

        # Avoid using global variables without explicit scoping
        PSAvoidGlobalVars = @{
            Enable = $true
        }

        # Review unused parameters in function declarations
        PSReviewUnusedParameter = @{
            Enable = $true
        }
    }
}
