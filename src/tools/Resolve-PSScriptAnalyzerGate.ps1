function ConvertTo-GitHubAnnotationField {
    # .SYNOPSIS
    # Escapes a value for use in a GitHub Actions annotation field.
    #
    # .DESCRIPTION
    # Converts a value to a string and applies GitHub Actions workflow command
    # escaping for annotation field values. Field values require the common
    # command escaping plus colon and comma escaping because those characters
    # delimit annotation metadata.
    #
    # .PARAMETER Value
    # The value to escape. Null values are emitted as an empty string.
    #
    # .EXAMPLE
    # ConvertTo-GitHubAnnotationField -Value 'C:\src\script.ps1'
    # # Returns C%3A\src\script.ps1
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The escaped annotation field value.
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
    [OutputType([string])]
    param(
        [AllowNull()]
        [object]$Value
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Value) {
        return ''
    }

    $strEscapedValue = [string]$Value
    $strEscapedValue = $strEscapedValue.Replace('%', '%25')
    $strEscapedValue = $strEscapedValue.Replace("`r", '%0D')
    $strEscapedValue = $strEscapedValue.Replace("`n", '%0A')
    $strEscapedValue = $strEscapedValue.Replace(':', '%3A')
    $strEscapedValue = $strEscapedValue.Replace(',', '%2C')

    return $strEscapedValue
}

function ConvertTo-GitHubAnnotationMessage {
    # .SYNOPSIS
    # Escapes a value for use as a GitHub Actions annotation message.
    #
    # .DESCRIPTION
    # Converts a value to a string and applies GitHub Actions workflow command
    # escaping for annotation message values. Message values do not require the
    # additional colon or comma escaping used by annotation metadata fields.
    #
    # .PARAMETER Value
    # The value to escape. Null values are emitted as an empty string.
    #
    # .EXAMPLE
    # ConvertTo-GitHubAnnotationMessage -Value "Line 1`nLine 2"
    # # Returns Line 1%0ALine 2
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The escaped annotation message.
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
    [OutputType([string])]
    param(
        [AllowNull()]
        [object]$Value
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Value) {
        return ''
    }

    $strEscapedValue = [string]$Value
    $strEscapedValue = $strEscapedValue.Replace('%', '%25')
    $strEscapedValue = $strEscapedValue.Replace("`r", '%0D')
    $strEscapedValue = $strEscapedValue.Replace("`n", '%0A')

    return $strEscapedValue
}

function ConvertTo-AzurePipelinesLoggingCommandPropertyValue {
    # .SYNOPSIS
    # Escapes a value for an Azure Pipelines logging-command property.
    #
    # .DESCRIPTION
    # Converts a value to a string and applies Azure Pipelines logging-command
    # escaping for property values inside the command metadata block. Property
    # values are separated by semicolons and terminated by a closing bracket, so
    # semicolons and closing brackets are escaped in addition to percent signs
    # and line endings.
    #
    # .PARAMETER Value
    # The value to escape. Null values are emitted as an empty string.
    #
    # .EXAMPLE
    # ConvertTo-AzurePipelinesLoggingCommandPropertyValue -Value 'PSRule;Name]'
    # # Returns PSRule%3BName%5D
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The escaped logging-command property value.
    #
    # .NOTES
    # PRIVATE/INTERNAL HELPER - This function is not part of the public
    # API surface. Parameters, return shape, and positional contract may
    # change without notice.
    #
    # Version: 1.0.20260619.0
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([string])]
    param(
        [AllowNull()]
        [object]$Value
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Value) {
        return ''
    }

    $strEscapedValue = [string]$Value
    $strEscapedValue = $strEscapedValue.Replace('%', '%AZP25')
    $strEscapedValue = $strEscapedValue.Replace("`r", '%0D')
    $strEscapedValue = $strEscapedValue.Replace("`n", '%0A')
    $strEscapedValue = $strEscapedValue.Replace(']', '%5D')
    $strEscapedValue = $strEscapedValue.Replace(';', '%3B')

    return $strEscapedValue
}

function ConvertTo-AzurePipelinesLoggingCommandMessage {
    # .SYNOPSIS
    # Escapes a value for an Azure Pipelines logging-command message.
    #
    # .DESCRIPTION
    # Converts a value to a string and applies Azure Pipelines logging-command
    # escaping for the free-text message after the command metadata block.
    # Message values are read to the end of the output line, so semicolons and
    # closing brackets are preserved while percent signs and line endings are
    # encoded to keep each rendered finding on one physical line.
    #
    # .PARAMETER Value
    # The value to escape. Null values are emitted as an empty string.
    #
    # .EXAMPLE
    # ConvertTo-AzurePipelinesLoggingCommandMessage -Value 'Use ]; keep 100%'
    # # Returns Use ]; keep 100%AZP25
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The escaped logging-command message value.
    #
    # .NOTES
    # PRIVATE/INTERNAL HELPER - This function is not part of the public
    # API surface. Parameters, return shape, and positional contract may
    # change without notice.
    #
    # Version: 1.0.20260619.0
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([string])]
    param(
        [AllowNull()]
        [object]$Value
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Value) {
        return ''
    }

    $strEscapedValue = [string]$Value
    $strEscapedValue = $strEscapedValue.Replace('%', '%AZP25')
    $strEscapedValue = $strEscapedValue.Replace("`r", '%0D')
    $strEscapedValue = $strEscapedValue.Replace("`n", '%0A')

    return $strEscapedValue
}

function Get-PSScriptAnalyzerFindingProperty {
    # .SYNOPSIS
    # Reads a named property from a PSScriptAnalyzer finding.
    #
    # .DESCRIPTION
    # Safely retrieves a property from an analyzer finding object. Missing
    # properties and null findings return null so callers can treat malformed
    # synthetic or analyzer output as unknown-severity findings.
    #
    # .PARAMETER Finding
    # The analyzer finding object to inspect.
    #
    # .PARAMETER Name
    # The property name to read from the finding.
    #
    # .EXAMPLE
    # Get-PSScriptAnalyzerFindingProperty -Finding $objFinding -Name 'Severity'
    # # Returns the Severity property value when present.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [object] The property value, or null when the property is absent.
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
    [OutputType([object])]
    param(
        [AllowNull()]
        [object]$Finding,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Finding) {
        return $null
    }

    $objProperty = $Finding.PSObject.Properties[$Name]
    if ($null -eq $objProperty) {
        return $null
    }

    return $objProperty.Value
}

function ConvertTo-PSScriptAnalyzerPositiveInteger {
    # .SYNOPSIS
    # Converts a value to a positive integer for annotation metadata.
    #
    # .DESCRIPTION
    # Converts line and column values from analyzer findings to positive
    # integers. Missing, non-numeric, zero, and negative values return 0 so the
    # caller can omit invalid annotation metadata.
    #
    # .PARAMETER Value
    # The value to parse as a positive integer.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerPositiveInteger -Value '12'
    # # Returns 12
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [int] A positive integer, or 0 when the value is not a positive integer.
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
    [OutputType([int])]
    param(
        [AllowNull()]
        [object]$Value
    )

    Set-StrictMode -Version Latest

    $intParsedValue = 0
    if ($null -eq $Value) {
        return $intParsedValue
    }

    [void]([int]::TryParse(([string]$Value), [ref]$intParsedValue))
    if ($intParsedValue -lt 1) {
        return 0
    }

    return $intParsedValue
}

function ConvertTo-RepositoryRelativePath {
    # .SYNOPSIS
    # Converts an absolute path to a repository-relative path.
    #
    # .DESCRIPTION
    # When a repository root is provided and the path is located beneath it, the
    # leading root prefix is removed so GitHub Actions annotations reference a
    # repository-relative file. Paths that are empty, lack a root, or fall
    # outside the root are returned unchanged. Backslashes are normalized to
    # forward slashes for the prefix comparison and in the repository-relative
    # result; a path returned unchanged keeps its original separators.
    #
    # .PARAMETER Path
    # The path to convert. A null path is returned as an empty string;
    # an empty or whitespace-only path is returned unchanged.
    #
    # .PARAMETER RepositoryRoot
    # The repository root to make the path relative to. Null and empty roots
    # leave the path unchanged.
    #
    # .EXAMPLE
    # ConvertTo-RepositoryRelativePath -Path '/runner/repo/src/a.ps1' -RepositoryRoot '/runner/repo'
    # # Returns src/a.ps1
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The repository-relative path, or the original path when it cannot
    # be made relative.
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
    [OutputType([string])]
    param(
        [AllowNull()]
        [object]$Path,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$RepositoryRoot
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Path) {
        return ''
    }

    $strPath = [string]$Path
    if ([string]::IsNullOrWhiteSpace($strPath)) {
        return $strPath
    }
    if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
        return $strPath
    }

    $strNormalizedPath = $strPath.Replace('\', '/')
    $strNormalizedRoot = $RepositoryRoot.Replace('\', '/').TrimEnd('/')
    if ($strNormalizedRoot.Length -eq 0) {
        return $strPath
    }

    $strRootPrefix = $strNormalizedRoot + '/'
    if ($strNormalizedPath.StartsWith($strRootPrefix, [System.StringComparison]::Ordinal)) {
        return $strNormalizedPath.Substring($strRootPrefix.Length)
    }

    return $strPath
}

function Test-PSScriptAnalyzerCiIndicator {
    # .SYNOPSIS
    # Tests whether a CI host indicator environment value is enabled.
    #
    # .DESCRIPTION
    # Treats the string value true, case-insensitively, as an enabled CI host
    # indicator. Empty, missing, false, and other values are treated as not
    # indicating that host so explicit false-like values do not trigger
    # annotation rendering.
    #
    # .PARAMETER Value
    # The environment value to test.
    #
    # .EXAMPLE
    # Test-PSScriptAnalyzerCiIndicator -Value 'True'
    # # Returns $true
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [bool] True when the value indicates an active CI host; otherwise false.
    #
    # .NOTES
    # PRIVATE/INTERNAL HELPER - This function is not part of the public
    # API surface. Parameters, return shape, and positional contract may
    # change without notice.
    #
    # Version: 1.0.20260619.0
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([bool])]
    param(
        [AllowNull()]
        [object]$Value
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Value) {
        return $false
    }

    return (([string]$Value).Trim() -eq 'true')
}

function Test-PSScriptAnalyzerRootedPath {
    # .SYNOPSIS
    # Tests whether a path is rooted for CI annotation purposes.
    #
    # .DESCRIPTION
    # Detects common rooted source path shapes without relying on the current
    # operating system's path parser. Azure Pipelines can run on Windows,
    # Linux, or macOS, while tests may need to reason about paths from any of
    # those agents.
    #
    # .PARAMETER Path
    # The path value to inspect.
    #
    # .EXAMPLE
    # Test-PSScriptAnalyzerRootedPath -Path '/agent/_work/1/s/file.ps1'
    # # Returns $true
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [bool] True when the path appears rooted; otherwise false.
    #
    # .NOTES
    # PRIVATE/INTERNAL HELPER - This function is not part of the public
    # API surface. Parameters, return shape, and positional contract may
    # change without notice.
    #
    # Version: 1.0.20260619.0
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([bool])]
    param(
        [AllowNull()]
        [object]$Path
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Path) {
        return $false
    }

    $strPath = [string]$Path
    if ([string]::IsNullOrWhiteSpace($strPath)) {
        return $false
    }

    return (
        ($strPath -match '^[A-Za-z]:[\\/]') -or
        ($strPath -match '^/') -or
        ($strPath -match '^\\\\')
    )
}

function ConvertTo-AzurePipelinesSourcePath {
    # .SYNOPSIS
    # Converts an analyzer path to an Azure Pipelines source path.
    #
    # .DESCRIPTION
    # Preserves rooted source paths and combines relative source paths with the
    # repository root when one is available. Azure Pipelines logging commands
    # attach file locations most reliably when source paths are absolute.
    #
    # .PARAMETER Path
    # The analyzer source path to convert.
    #
    # .PARAMETER RepositoryRoot
    # The repository root used to make relative paths absolute.
    #
    # .EXAMPLE
    # ConvertTo-AzurePipelinesSourcePath -Path 'src/a.ps1' -RepositoryRoot '/agent/_work/1/s'
    # # Returns /agent/_work/1/s/src/a.ps1
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The Azure Pipelines source path.
    #
    # .NOTES
    # PRIVATE/INTERNAL HELPER - This function is not part of the public
    # API surface. Parameters, return shape, and positional contract may
    # change without notice.
    #
    # Version: 1.0.20260619.0
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([string])]
    param(
        [AllowNull()]
        [object]$Path,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$RepositoryRoot
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Path) {
        return ''
    }

    $strPath = [string]$Path
    if ([string]::IsNullOrWhiteSpace($strPath)) {
        return $strPath
    }

    if (Test-PSScriptAnalyzerRootedPath -Path $strPath) {
        return $strPath
    }

    if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
        return $strPath
    }

    $arrSeparator = [char[]]@('/', '\')
    $strTrimmedRoot = $RepositoryRoot.TrimEnd($arrSeparator)
    if ([string]::IsNullOrWhiteSpace($strTrimmedRoot)) {
        return $strPath
    }

    $strSeparator = '/'
    if (($strTrimmedRoot.Contains('\')) -and (-not $strTrimmedRoot.Contains('/'))) {
        $strSeparator = '\'
    }

    $strTrimmedPath = $strPath.TrimStart($arrSeparator)

    return ('{0}{1}{2}' -f $strTrimmedRoot, $strSeparator, $strTrimmedPath)
}

function ConvertTo-PSScriptAnalyzerSingleLineText {
    # .SYNOPSIS
    # Converts text to a single physical output line.
    #
    # .DESCRIPTION
    # Converts a value to a string and replaces carriage returns and newlines
    # with spaces so plain console rendering emits one output line per finding.
    #
    # .PARAMETER Value
    # The value to convert. Null values are emitted as an empty string.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerSingleLineText -Value "Line 1`nLine 2"
    # # Returns Line 1 Line 2
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The single-line text value.
    #
    # .NOTES
    # PRIVATE/INTERNAL HELPER - This function is not part of the public
    # API surface. Parameters, return shape, and positional contract may
    # change without notice.
    #
    # Version: 1.0.20260619.0
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([string])]
    param(
        [AllowNull()]
        [object]$Value
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Value) {
        return ''
    }

    $strSingleLineValue = [string]$Value
    $strSingleLineValue = $strSingleLineValue.Replace("`r`n", ' ')
    $strSingleLineValue = $strSingleLineValue.Replace("`r", ' ')
    $strSingleLineValue = $strSingleLineValue.Replace("`n", ' ')

    return $strSingleLineValue
}

function ConvertTo-PSScriptAnalyzerPlainTextLine {
    # .SYNOPSIS
    # Renders a normalized finding as a plain console line.
    #
    # .DESCRIPTION
    # Produces a deterministic, single-line text representation of a normalized
    # PSScriptAnalyzer finding for local runs and informational Azure Pipelines
    # diagnostics that should not be promoted to warnings or errors.
    #
    # .PARAMETER Finding
    # The normalized finding to render.
    #
    # .PARAMETER Path
    # The path to display for the finding. When omitted or empty, the finding's
    # normalized ScriptPath is used.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerPlainTextLine -Finding $objFinding
    # # Returns a single-line diagnostic.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The plain diagnostic line.
    #
    # .NOTES
    # PRIVATE/INTERNAL HELPER - This function is not part of the public
    # API surface. Parameters, return shape, and positional contract may
    # change without notice.
    #
    # Version: 1.0.20260619.0
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$Finding,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$Path
    )

    Set-StrictMode -Version Latest

    $strDisplayPath = $Path
    if ([string]::IsNullOrWhiteSpace($strDisplayPath)) {
        $strDisplayPath = [string]$Finding.ScriptPath
    }
    if ([string]::IsNullOrWhiteSpace($strDisplayPath)) {
        $strDisplayPath = '<unknown>'
    }

    $strLocation = ConvertTo-PSScriptAnalyzerSingleLineText -Value $strDisplayPath
    if ([int]$Finding.Line -gt 0) {
        $strLocation = '{0}:{1}' -f $strLocation, [int]$Finding.Line
        if ([int]$Finding.Column -gt 0) {
            $strLocation = '{0}:{1}' -f $strLocation, [int]$Finding.Column
        }
    }

    $strRuleName = ConvertTo-PSScriptAnalyzerSingleLineText -Value $Finding.RuleName
    $strMessage = ConvertTo-PSScriptAnalyzerSingleLineText -Value $Finding.Message

    return ('PSScriptAnalyzer {0}: {1} [{2}] {3}' -f $Finding.DisplaySeverity, $strLocation, $strRuleName, $strMessage)
}

function ConvertTo-PSScriptAnalyzerGitHubAnnotationCommand {
    # .SYNOPSIS
    # Renders a normalized finding as a GitHub Actions annotation command.
    #
    # .DESCRIPTION
    # Converts a normalized PSScriptAnalyzer finding into the GitHub Actions
    # workflow command syntax used by the existing PowerShell CI workflow.
    #
    # .PARAMETER Finding
    # The normalized finding to render.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerGitHubAnnotationCommand -Finding $objFinding
    # # Returns a ::warning file=...::... annotation command.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The GitHub Actions annotation command.
    #
    # .NOTES
    # PRIVATE/INTERNAL HELPER - This function is not part of the public
    # API surface. Parameters, return shape, and positional contract may
    # change without notice.
    #
    # Version: 1.0.20260619.0
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$Finding
    )

    Set-StrictMode -Version Latest

    $listAnnotationField = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($Finding.ScriptPath)) {
        $listAnnotationField.Add(('file={0}' -f (ConvertTo-GitHubAnnotationField -Value $Finding.ScriptPath)))
    }
    if ([int]$Finding.Line -gt 0) {
        $listAnnotationField.Add(('line={0}' -f [int]$Finding.Line))
    }
    if ([int]$Finding.Column -gt 0) {
        $listAnnotationField.Add(('col={0}' -f [int]$Finding.Column))
    }

    $strAnnotationMessage = '[{0}] {1} - {2}' -f $Finding.DisplaySeverity, $Finding.RuleName, $Finding.Message
    $strEscapedAnnotationMessage = ConvertTo-GitHubAnnotationMessage -Value $strAnnotationMessage
    if ($listAnnotationField.Count -gt 0) {
        $strAnnotationField = $listAnnotationField.ToArray() -join ','
        return ('::{0} {1}::{2}' -f $Finding.AnnotationLevel, $strAnnotationField, $strEscapedAnnotationMessage)
    }

    return ('::{0}::{1}' -f $Finding.AnnotationLevel, $strEscapedAnnotationMessage)
}

function ConvertTo-PSScriptAnalyzerAzurePipelinesOutputLine {
    # .SYNOPSIS
    # Renders a normalized finding for Azure Pipelines output.
    #
    # .DESCRIPTION
    # Converts a normalized PSScriptAnalyzer finding into an Azure Pipelines
    # task.logissue logging command for warning and error diagnostics. The
    # Information severity is rendered as a plain line so it remains visible
    # without becoming an Azure warning or error.
    #
    # .PARAMETER Finding
    # The normalized finding to render.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerAzurePipelinesOutputLine -Finding $objFinding
    # # Returns a ##vso[task.logissue ...] command for warnings and errors.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The Azure Pipelines output line.
    #
    # .NOTES
    # PRIVATE/INTERNAL HELPER - This function is not part of the public
    # API surface. Parameters, return shape, and positional contract may
    # change without notice.
    #
    # Version: 1.0.20260619.0
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$Finding
    )

    Set-StrictMode -Version Latest

    if ($Finding.Severity -eq 'Information') {
        return ConvertTo-PSScriptAnalyzerPlainTextLine -Finding $Finding -Path $Finding.AzureSourcePath
    }

    $strIssueType = 'error'
    if ($Finding.Severity -eq 'Warning') {
        $strIssueType = 'warning'
    }

    $listProperty = [System.Collections.Generic.List[string]]::new()
    $listProperty.Add(('type={0}' -f $strIssueType))
    if (-not [string]::IsNullOrWhiteSpace($Finding.AzureSourcePath)) {
        $strEscapedSourcePath = ConvertTo-AzurePipelinesLoggingCommandPropertyValue -Value $Finding.AzureSourcePath
        $listProperty.Add(('sourcepath={0}' -f $strEscapedSourcePath))
    }
    if ([int]$Finding.Line -gt 0) {
        $listProperty.Add(('linenumber={0}' -f [int]$Finding.Line))
    }
    if ([int]$Finding.Column -gt 0) {
        $listProperty.Add(('columnnumber={0}' -f [int]$Finding.Column))
    }
    if (-not [string]::IsNullOrWhiteSpace($Finding.RuleName)) {
        $strEscapedRuleName = ConvertTo-AzurePipelinesLoggingCommandPropertyValue -Value $Finding.RuleName
        $listProperty.Add(('code={0}' -f $strEscapedRuleName))
    }

    $strProperty = ($listProperty.ToArray() -join ';') + ';'
    $strMessage = '[{0}] {1} - {2}' -f $Finding.DisplaySeverity, $Finding.RuleName, $Finding.Message
    $strEscapedMessage = ConvertTo-AzurePipelinesLoggingCommandMessage -Value $strMessage

    return ('##vso[task.logissue {0}]{1}' -f $strProperty, $strEscapedMessage)
}

function ConvertTo-PSScriptAnalyzerOutputLine {
    # .SYNOPSIS
    # Renders normalized findings in the selected output format.
    #
    # .DESCRIPTION
    # Converts normalized PSScriptAnalyzer findings into GitHub Actions
    # annotations, Azure Pipelines logging commands, or plain console lines.
    #
    # .PARAMETER Finding
    # The normalized findings to render.
    #
    # .PARAMETER AnnotationFormat
    # The resolved annotation format to use.
    #
    # .EXAMPLE
    # $arrLine = @(ConvertTo-PSScriptAnalyzerOutputLine -Finding $arrFinding -AnnotationFormat 'Plain')
    # # Returns plain diagnostic lines, one streamed string per finding.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The rendered output lines, streamed one per finding.
    #
    # .NOTES
    # PRIVATE/INTERNAL HELPER - This function is not part of the public
    # API surface. Parameters, return shape, and positional contract may
    # change without notice.
    #
    # Version: 1.0.20260619.1
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([string])]
    param(
        [AllowNull()]
        [object[]]$Finding,

        [Parameter(Mandatory = $true)]
        [ValidateSet('GitHubActions', 'AzurePipelines', 'Plain')]
        [string]$AnnotationFormat
    )

    Set-StrictMode -Version Latest

    $arrFinding = @()
    if ($null -ne $Finding) {
        $arrFinding = @($Finding)
    }

    foreach ($objFinding in $arrFinding) {
        if ($AnnotationFormat -eq 'GitHubActions') {
            ConvertTo-PSScriptAnalyzerGitHubAnnotationCommand -Finding $objFinding
        } elseif ($AnnotationFormat -eq 'AzurePipelines') {
            ConvertTo-PSScriptAnalyzerAzurePipelinesOutputLine -Finding $objFinding
        } else {
            ConvertTo-PSScriptAnalyzerPlainTextLine -Finding $objFinding
        }
    }
}

function Resolve-PSScriptAnalyzerAnnotationFormat {
    # .SYNOPSIS
    # Resolves the requested analyzer annotation format.
    #
    # .DESCRIPTION
    # Resolves explicit annotation formats directly and auto-detects supported
    # CI hosts from environment indicators. Auto mode fails closed when both
    # GitHub Actions and Azure Pipelines indicators are present.
    #
    # .PARAMETER AnnotationFormat
    # The requested format: Auto, GitHubActions, AzurePipelines, or Plain.
    #
    # .EXAMPLE
    # Resolve-PSScriptAnalyzerAnnotationFormat -AnnotationFormat 'Auto'
    # # Returns the detected format, or Plain outside supported CI hosts.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The resolved annotation format.
    #
    # .NOTES
    # PRIVATE/INTERNAL HELPER - This function is not part of the public
    # API surface. Parameters, return shape, and positional contract may
    # change without notice.
    #
    # Version: 1.0.20260619.0
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([string])]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$AnnotationFormat
    )

    Set-StrictMode -Version Latest

    if (-not [string]::IsNullOrWhiteSpace($AnnotationFormat)) {
        switch ($AnnotationFormat.Trim().ToLowerInvariant()) {
            'githubactions' {
                return 'GitHubActions'
            }
            'azurepipelines' {
                return 'AzurePipelines'
            }
            'plain' {
                return 'Plain'
            }
        }
    }

    $boolGitHubActions = Test-PSScriptAnalyzerCiIndicator -Value $env:GITHUB_ACTIONS
    $boolAzurePipelines = Test-PSScriptAnalyzerCiIndicator -Value $env:TF_BUILD

    if (($boolGitHubActions) -and ($boolAzurePipelines)) {
        throw 'Both GITHUB_ACTIONS and TF_BUILD indicate supported CI hosts. Pass an explicit -AnnotationFormat value (GitHubActions, AzurePipelines, or Plain).'
    }

    if ($boolGitHubActions) {
        return 'GitHubActions'
    }

    if ($boolAzurePipelines) {
        return 'AzurePipelines'
    }

    return 'Plain'
}

function Resolve-PSScriptAnalyzerGate {
    # .SYNOPSIS
    # Resolves PSScriptAnalyzer findings into a CI gate decision.
    #
    # .DESCRIPTION
    # Resolves a configured PSScriptAnalyzer gate mode and evaluates analyzer
    # findings by severity. Strict mode fails Error, Warning, and unknown
    # severity findings. First-adoption mode fails Error and unknown severity
    # findings while annotating Warning and Information findings as tracked
    # adoption debt. Information findings are annotation-only in both modes.
    #
    # Missing, empty, and unrecognized mode values resolve to strict mode.
    # The returned object includes deterministic summary data, rendered output
    # lines for every finding, normalized finding records, and the final gate
    # decision.
    #
    # .PARAMETER Mode
    # The requested gate mode. Supported values are strict and first-adoption.
    #
    # .PARAMETER AnalyzerFinding
    # The PSScriptAnalyzer findings to evaluate.
    #
    # .PARAMETER RepositoryRoot
    # Optional repository root used to render annotation file paths relative to
    # the repository for GitHub Actions and to build absolute Azure Pipelines
    # source paths from relative analyzer paths. Defaults to the
    # GITHUB_WORKSPACE environment variable. When both RepositoryRoot and
    # GITHUB_WORKSPACE are empty, the BUILD_SOURCESDIRECTORY environment
    # variable is used as the effective repository root so Azure Pipelines
    # callers still resolve absolute source paths.
    #
    # .PARAMETER AnnotationFormat
    # The annotation/output format to render. Supported values are Auto,
    # GitHubActions, AzurePipelines, and Plain. Auto uses GitHub Actions when
    # GITHUB_ACTIONS indicates GitHub Actions, Azure Pipelines when TF_BUILD
    # indicates Azure Pipelines, Plain when no supported CI host is detected,
    # and fails closed if both supported CI indicators are present.
    #
    # .EXAMPLE
    # $objGate = Resolve-PSScriptAnalyzerGate -Mode 'first-adoption' -AnalyzerFinding $arrFinding
    # if ($objGate.ShouldFail) {
    #     exit 1
    # }
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [pscustomobject] Gate result with Mode, ShouldFail, Summary,
    # SummaryLines, AnnotationFormat, AnnotationCommands, Findings, RuleCounts,
    # FileCounts, TopRules, TopFiles, RecommendedMode, and IssueReadyMarkdown
    # properties.
    #
    # .NOTES
    # Version: 1.2.20260619.1
    # Positional parameters are not supported.
    #
    [CmdletBinding(PositionalBinding = $false)]
    [OutputType([pscustomobject])]
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Mode,

        [AllowNull()]
        [object[]]$AnalyzerFinding,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$RepositoryRoot = $env:GITHUB_WORKSPACE,

        [ValidateSet('Auto', 'GitHubActions', 'AzurePipelines', 'Plain')]
        [string]$AnnotationFormat = 'Auto'
    )

    Set-StrictMode -Version Latest

    $strRequestedMode = ''
    if ($null -ne $Mode) {
        $strRequestedMode = $Mode.Trim().ToLowerInvariant()
    }

    $strResolvedMode = 'strict'
    if ($strRequestedMode -eq 'first-adoption') {
        $strResolvedMode = 'first-adoption'
    }

    $strResolvedAnnotationFormat = Resolve-PSScriptAnalyzerAnnotationFormat -AnnotationFormat $AnnotationFormat

    $strEffectiveRepositoryRoot = $RepositoryRoot
    if (
        [string]::IsNullOrWhiteSpace($strEffectiveRepositoryRoot) -and
        (-not [string]::IsNullOrWhiteSpace($env:BUILD_SOURCESDIRECTORY))
    ) {
        $strEffectiveRepositoryRoot = $env:BUILD_SOURCESDIRECTORY
    }

    $arrAnalyzerFinding = @()
    if ($null -ne $AnalyzerFinding) {
        $arrAnalyzerFinding = @($AnalyzerFinding)
    }

    $listNormalizedFinding = [System.Collections.Generic.List[pscustomobject]]::new()

    $intErrorCount = 0
    $intWarningCount = 0
    $intInformationCount = 0
    $intUnknownSeverityCount = 0
    $intBlockingCount = 0
    $intDebtCount = 0
    $hashtableRuleCount = @{}
    $hashtableFileCount = @{}

    foreach ($objFinding in $arrAnalyzerFinding) {
        $objSeverityValue = Get-PSScriptAnalyzerFindingProperty -Finding $objFinding -Name 'Severity'
        $strOriginalSeverity = ''
        if ($null -ne $objSeverityValue) {
            $strOriginalSeverity = ([string]$objSeverityValue).Trim()
        }

        $strNormalizedSeverity = 'Unknown'
        switch ($strOriginalSeverity.ToLowerInvariant()) {
            'error' {
                $strNormalizedSeverity = 'Error'
                $intErrorCount++
                break
            }
            'warning' {
                $strNormalizedSeverity = 'Warning'
                $intWarningCount++
                break
            }
            'information' {
                $strNormalizedSeverity = 'Information'
                $intInformationCount++
                break
            }
            default {
                $intUnknownSeverityCount++
                break
            }
        }

        $boolFailsGate = $false
        if ($strNormalizedSeverity -eq 'Error') {
            $boolFailsGate = $true
        } elseif ($strNormalizedSeverity -eq 'Warning') {
            $boolFailsGate = ($strResolvedMode -eq 'strict')
        } elseif ($strNormalizedSeverity -eq 'Unknown') {
            $boolFailsGate = $true
        }

        $boolTrackedDebt = (
            ($strResolvedMode -eq 'first-adoption') -and
            (-not $boolFailsGate) -and
            (($strNormalizedSeverity -eq 'Warning') -or ($strNormalizedSeverity -eq 'Information'))
        )

        if ($boolFailsGate) {
            $intBlockingCount++
        }

        if ($boolTrackedDebt) {
            $intDebtCount++
        }

        $strAnnotationLevel = 'notice'
        if (($strNormalizedSeverity -eq 'Error') -or ($strNormalizedSeverity -eq 'Unknown')) {
            $strAnnotationLevel = 'error'
        } elseif ($strNormalizedSeverity -eq 'Warning') {
            $strAnnotationLevel = 'warning'
        }

        $strRuleName = [string](Get-PSScriptAnalyzerFindingProperty -Finding $objFinding -Name 'RuleName')
        if ([string]::IsNullOrWhiteSpace($strRuleName)) {
            $strRuleName = 'PSScriptAnalyzer'
        }

        $strMessage = [string](Get-PSScriptAnalyzerFindingProperty -Finding $objFinding -Name 'Message')
        if ([string]::IsNullOrWhiteSpace($strMessage)) {
            $strMessage = 'PSScriptAnalyzer finding did not include a message.'
        }

        $strScriptPath = [string](Get-PSScriptAnalyzerFindingProperty -Finding $objFinding -Name 'ScriptPath')
        if ([string]::IsNullOrWhiteSpace($strScriptPath)) {
            $strScriptPath = [string](Get-PSScriptAnalyzerFindingProperty -Finding $objFinding -Name 'FileName')
        }
        $strSourcePath = $strScriptPath
        $strAzureSourcePath = ConvertTo-AzurePipelinesSourcePath -Path $strSourcePath -RepositoryRoot $strEffectiveRepositoryRoot
        $strScriptPath = ConvertTo-RepositoryRelativePath -Path $strScriptPath -RepositoryRoot $strEffectiveRepositoryRoot
        $strCountedScriptPath = $strScriptPath
        if ([string]::IsNullOrWhiteSpace($strCountedScriptPath)) {
            $strCountedScriptPath = '<unknown>'
        }

        if ($hashtableRuleCount.ContainsKey($strRuleName)) {
            $hashtableRuleCount[$strRuleName] = [int]$hashtableRuleCount[$strRuleName] + 1
        } else {
            $hashtableRuleCount[$strRuleName] = 1
        }

        if ($hashtableFileCount.ContainsKey($strCountedScriptPath)) {
            $hashtableFileCount[$strCountedScriptPath] = [int]$hashtableFileCount[$strCountedScriptPath] + 1
        } else {
            $hashtableFileCount[$strCountedScriptPath] = 1
        }

        $intLine = ConvertTo-PSScriptAnalyzerPositiveInteger -Value (
            Get-PSScriptAnalyzerFindingProperty -Finding $objFinding -Name 'Line'
        )
        $intColumn = ConvertTo-PSScriptAnalyzerPositiveInteger -Value (
            Get-PSScriptAnalyzerFindingProperty -Finding $objFinding -Name 'Column'
        )

        $strDisplaySeverity = $strNormalizedSeverity
        if (($strNormalizedSeverity -eq 'Unknown') -and (-not [string]::IsNullOrWhiteSpace($strOriginalSeverity))) {
            $strDisplaySeverity = "Unknown ($strOriginalSeverity)"
        }

        $listNormalizedFinding.Add(
            [pscustomobject]@{
                Severity = $strNormalizedSeverity
                OriginalSeverity = $strOriginalSeverity
                DisplaySeverity = $strDisplaySeverity
                RuleName = $strRuleName
                Message = $strMessage
                ScriptPath = $strScriptPath
                SourcePath = $strSourcePath
                AzureSourcePath = $strAzureSourcePath
                Line = $intLine
                Column = $intColumn
                AnnotationLevel = $strAnnotationLevel
                FailsGate = $boolFailsGate
                TrackedDebt = $boolTrackedDebt
            }
        )
    }

    $intTotalCount = $listNormalizedFinding.Count
    $boolShouldFail = ($intBlockingCount -gt 0)
    $strRecommendedMode = 'strict'
    if (($intErrorCount -eq 0) -and ($intUnknownSeverityCount -eq 0) -and (($intWarningCount + $intInformationCount) -gt 0)) {
        $strRecommendedMode = 'first-adoption'
    }

    $arrRuleCount = @(
        foreach ($strRule in $hashtableRuleCount.Keys) {
            [pscustomobject]@{
                RuleName = [string]$strRule
                Count = [int]$hashtableRuleCount[$strRule]
            }
        }
    )
    $arrRuleSortProperty = @(
        @{ Expression = 'Count'; Descending = $true }
        @{ Expression = 'RuleName'; Descending = $false }
    )
    $arrRuleCount = @(
        $arrRuleCount | Sort-Object -Property $arrRuleSortProperty
    )

    $arrFileCount = @(
        foreach ($strFile in $hashtableFileCount.Keys) {
            [pscustomobject]@{
                ScriptPath = [string]$strFile
                Count = [int]$hashtableFileCount[$strFile]
            }
        }
    )
    $arrFileSortProperty = @(
        @{ Expression = 'Count'; Descending = $true }
        @{ Expression = 'ScriptPath'; Descending = $false }
    )
    $arrFileCount = @(
        $arrFileCount | Sort-Object -Property $arrFileSortProperty
    )

    $arrTopRule = @($arrRuleCount | Select-Object -First 5)
    $arrTopFile = @($arrFileCount | Select-Object -First 5)

    $objSummary = [pscustomobject]@{
        TotalCount = $intTotalCount
        ErrorCount = $intErrorCount
        WarningCount = $intWarningCount
        InformationCount = $intInformationCount
        UnknownSeverityCount = $intUnknownSeverityCount
        BlockingCount = $intBlockingCount
        DebtCount = $intDebtCount
        RecommendedMode = $strRecommendedMode
    }

    $listSummaryLine = [System.Collections.Generic.List[string]]::new()
    $listSummaryLine.Add(('PSScriptAnalyzer gate mode: {0}.' -f $strResolvedMode))

    if ($intTotalCount -eq 0) {
        $listSummaryLine.Add('No PSScriptAnalyzer findings were reported.')
        $listSummaryLine.Add('Recommended gate mode: strict.')
        $listSummaryLine.Add('Result: pass.')
    } else {
        $strFindingSummary = 'Findings: {0} total; {1} Error; {2} Warning; {3} Information; {4} unknown severity.' -f $intTotalCount, $intErrorCount, $intWarningCount, $intInformationCount, $intUnknownSeverityCount
        $listSummaryLine.Add($strFindingSummary)
        $listSummaryLine.Add(('Recommended gate mode from full scan: {0}.' -f $strRecommendedMode))

        if ($arrTopRule.Count -gt 0) {
            $strTopRuleSummary = (
                $arrTopRule |
                    ForEach-Object { '{0} ({1})' -f $_.RuleName, $_.Count }
            ) -join '; '
            $listSummaryLine.Add(('Top rule findings: {0}.' -f $strTopRuleSummary))
        }

        if ($arrTopFile.Count -gt 0) {
            $strTopFileSummary = (
                $arrTopFile |
                    ForEach-Object { '{0} ({1})' -f $_.ScriptPath, $_.Count }
            ) -join '; '
            $listSummaryLine.Add(('Top files by findings: {0}.' -f $strTopFileSummary))
        }

        if ($strResolvedMode -eq 'first-adoption') {
            if ($intDebtCount -gt 0) {
                $strWarningDebtSummary = 'Warning debt: {0} Warning and {1} Information finding(s) are annotated as tracked debt.' -f $intWarningCount, $intInformationCount
                $listSummaryLine.Add($strWarningDebtSummary)
                $listSummaryLine.Add(
                    'Record warning debt in _TODO-repo-init.md or a post-adoption issue, then return PSSCRIPTANALYZER_GATE_MODE to strict after the debt is remediated.'
                )
            } else {
                $listSummaryLine.Add('Warning debt: none.')
            }

            if ($boolShouldFail) {
                $listSummaryLine.Add('Result: fail because Error or unknown-severity findings are present.')
            } else {
                $listSummaryLine.Add('Result: pass; Warning and Information findings were annotated without failing first-adoption mode.')
            }
        } else {
            $listSummaryLine.Add('Strict gate: Error, Warning, and unknown-severity findings fail CI; Information findings are annotation-only.')
            if ($boolShouldFail) {
                $listSummaryLine.Add('Result: fail because blocking findings are present.')
            } else {
                $listSummaryLine.Add('Result: pass.')
            }
        }
    }

    $listIssueMarkdown = [System.Collections.Generic.List[string]]::new()
    if ($strRecommendedMode -eq 'first-adoption') {
        $listIssueMarkdown.Add('## PSScriptAnalyzer First-Adoption Debt Cleanup')
        $listIssueMarkdown.Add('')
        $listIssueMarkdown.Add('### Summary')
        $listIssueMarkdown.Add(('- Findings: {0} total; {1} Warning; {2} Information.' -f $intTotalCount, $intWarningCount, $intInformationCount))
        $listIssueMarkdown.Add('- Recommended temporary gate mode: `first-adoption`.')
        $listIssueMarkdown.Add('- Target final gate mode after cleanup: `strict`.')
        if ($arrTopRule.Count -gt 0) {
            $listIssueMarkdown.Add('')
            $listIssueMarkdown.Add('### Top Rules')
            foreach ($objRuleCount in $arrTopRule) {
                $listIssueMarkdown.Add(('- `{0}`: {1}' -f $objRuleCount.RuleName, $objRuleCount.Count))
            }
        }
        if ($arrTopFile.Count -gt 0) {
            $listIssueMarkdown.Add('')
            $listIssueMarkdown.Add('### Top Files')
            foreach ($objFileCount in $arrTopFile) {
                $listIssueMarkdown.Add(('- `{0}`: {1}' -f $objFileCount.ScriptPath, $objFileCount.Count))
            }
        }
        $listIssueMarkdown.Add('')
        $listIssueMarkdown.Add('### Acceptance Criteria')
        $listIssueMarkdown.Add('- PSScriptAnalyzer Warning and Information findings are remediated or intentionally suppressed with narrow justification.')
        $listIssueMarkdown.Add('- `PSSCRIPTANALYZER_GATE_MODE` is returned to `strict` after cleanup.')
    }

    $arrAnnotationCommand = @(
        ConvertTo-PSScriptAnalyzerOutputLine `
            -Finding ([object[]]$listNormalizedFinding.ToArray()) `
            -AnnotationFormat $strResolvedAnnotationFormat
    )

    return [pscustomobject]@{
        Mode = $strResolvedMode
        ShouldFail = $boolShouldFail
        Summary = $objSummary
        SummaryLines = [string[]]$listSummaryLine.ToArray()
        AnnotationFormat = $strResolvedAnnotationFormat
        AnnotationCommands = [string[]]$arrAnnotationCommand
        Findings = [object[]]$listNormalizedFinding.ToArray()
        RuleCounts = [object[]]$arrRuleCount
        FileCounts = [object[]]$arrFileCount
        TopRules = [object[]]$arrTopRule
        TopFiles = [object[]]$arrTopFile
        RecommendedMode = $strRecommendedMode
        IssueReadyMarkdown = [string[]]$listIssueMarkdown.ToArray()
    }
}
