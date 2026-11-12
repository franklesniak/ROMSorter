function ConvertTo-PSScriptAnalyzerCandidateSingleLineText {
    # .SYNOPSIS
    # Converts candidate text to one physical output line.
    #
    # .DESCRIPTION
    # Converts a value to a string and escapes carriage returns and newlines so
    # candidate diagnostics can be rendered safely in console logs, Markdown,
    # and CI command payloads without changing the structured candidate data.
    #
    # .PARAMETER Value
    # The value to render. Null values are emitted as an empty string.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerCandidateSingleLineText -Value "a`nb"
    #
    # # Returns a\nb.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] A single-line rendering of the input value.
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
    param(
        [AllowNull()]
        [object]$Value
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Value) {
        return ''
    }

    $strValue = [string]$Value
    $strValue = $strValue.Replace("`r", '\r')
    $strValue = $strValue.Replace("`n", '\n')

    return $strValue
}

function ConvertTo-PSScriptAnalyzerCandidateGitHubCommandValue {
    # .SYNOPSIS
    # Escapes a candidate value for GitHub Actions command output.
    #
    # .DESCRIPTION
    # Applies the GitHub Actions workflow-command escaping required for a value
    # rendered inside an annotation command message.
    #
    # .PARAMETER Value
    # The value to escape. Null values are emitted as an empty string.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerCandidateGitHubCommandValue -Value "a`nb"
    #
    # # Returns a%0Ab.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The escaped command value.
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

function ConvertTo-PSScriptAnalyzerCandidateAzureCommandPropertyValue {
    # .SYNOPSIS
    # Escapes a candidate value for an Azure Pipelines command property.
    #
    # .DESCRIPTION
    # Applies Azure Pipelines logging-command escaping for property values that
    # appear inside the command metadata block.
    #
    # .PARAMETER Value
    # The value to escape. Null values are emitted as an empty string.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerCandidateAzureCommandPropertyValue -Value 'a;b]'
    #
    # # Returns a%3Bb%5D.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The escaped command property value.
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

function ConvertTo-PSScriptAnalyzerCandidateAzureCommandMessage {
    # .SYNOPSIS
    # Escapes a candidate value for an Azure Pipelines command message.
    #
    # .DESCRIPTION
    # Applies Azure Pipelines logging-command escaping for the free-text
    # message portion after the command metadata block.
    #
    # .PARAMETER Value
    # The value to escape. Null values are emitted as an empty string.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerCandidateAzureCommandMessage -Value "a`nb"
    #
    # # Returns a%0Ab.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The escaped command message.
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

function Get-PSScriptAnalyzerCandidateAllowedExtension {
    # .SYNOPSIS
    # Gets the effective analyzer candidate extension set.
    #
    # .DESCRIPTION
    # Returns the extension set used for both candidate selection and resolved
    # reparse-target validation. PSScriptAnalyzer supports PowerShell scripts,
    # modules, and data files.
    #
    # .EXAMPLE
    # Get-PSScriptAnalyzerCandidateAllowedExtension
    #
    # # Returns .ps1, .psm1, and .psd1.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string[]] The analyzer candidate extensions, lower-case with a leading dot.
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
    [OutputType([string[]])]
    param()

    Set-StrictMode -Version Latest

    return [string[]]@('.ps1', '.psm1', '.psd1')
}

function Test-PSScriptAnalyzerCandidateExtension {
    # .SYNOPSIS
    # Tests whether a path has an allowed analyzer extension.
    #
    # .DESCRIPTION
    # Compares the path extension against the shared candidate extension set
    # using an ordinal case-insensitive comparison.
    #
    # .PARAMETER Path
    # The path to inspect.
    #
    # .EXAMPLE
    # Test-PSScriptAnalyzerCandidateExtension -Path 'tools/Build.ps1'
    #
    # # Returns true.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [bool] True when the path is in analyzer scope; otherwise false.
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
    [OutputType([bool])]
    param(
        [AllowNull()]
        [object]$Path
    )

    Set-StrictMode -Version Latest

    if ($null -eq $Path) {
        return $false
    }

    $strExtension = [System.IO.Path]::GetExtension([string]$Path)
    $setAllowedExtension = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($strAllowedExtension in Get-PSScriptAnalyzerCandidateAllowedExtension) {
        [void]($setAllowedExtension.Add($strAllowedExtension))
    }

    return $setAllowedExtension.Contains($strExtension)
}

function Test-PSScriptAnalyzerExactPathSegment {
    # .SYNOPSIS
    # Tests whether a repository-relative path contains an exact segment.
    #
    # .DESCRIPTION
    # Normalizes path separators to forward slashes, splits the path into
    # segments, and compares each segment to the requested segment name using an
    # ordinal case-insensitive comparison. Substring matches are intentionally
    # not treated as matches.
    #
    # .PARAMETER RepositoryRelativePath
    # The repository-relative path to inspect.
    #
    # .PARAMETER SegmentName
    # The exact segment name to match.
    #
    # .EXAMPLE
    # Test-PSScriptAnalyzerExactPathSegment -RepositoryRelativePath 'node_modules/.bin/tool.ps1' -SegmentName 'node_modules'
    #
    # # Returns true.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [bool] True when the exact segment is present; otherwise false.
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
    [OutputType([bool])]
    param(
        [AllowNull()]
        [object]$RepositoryRelativePath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SegmentName
    )

    Set-StrictMode -Version Latest

    if ($null -eq $RepositoryRelativePath) {
        return $false
    }

    $arrSegment = ([string]$RepositoryRelativePath).Replace('\', '/').Split([char[]]@('/'))
    foreach ($strSegment in $arrSegment) {
        if (
            [string]::Equals(
                $strSegment,
                $SegmentName,
                [System.StringComparison]::OrdinalIgnoreCase
            )
        ) {
            return $true
        }
    }

    return $false
}

function Test-PSScriptAnalyzerNodeModuleSegment {
    # .SYNOPSIS
    # Tests whether a repository-relative path contains a node_modules segment.
    #
    # .DESCRIPTION
    # Compares normalized repository-relative path segments against the exact
    # dependency-shim segment name. Substring matches such as
    # node_modules_helper are intentionally not excluded.
    #
    # .PARAMETER RepositoryRelativePath
    # The repository-relative path to inspect.
    #
    # .EXAMPLE
    # Test-PSScriptAnalyzerNodeModuleSegment -RepositoryRelativePath 'node_modules/.bin/tool.ps1'
    #
    # # Returns true.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [bool] True when an exact node_modules path segment is present.
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
    [OutputType([bool])]
    param(
        [AllowNull()]
        [object]$RepositoryRelativePath
    )

    Set-StrictMode -Version Latest

    if ($null -eq $RepositoryRelativePath) {
        return $false
    }

    return Test-PSScriptAnalyzerExactPathSegment `
        -RepositoryRelativePath $RepositoryRelativePath `
        -SegmentName 'node_modules'
}

function Test-PSScriptAnalyzerGitDirectorySegment {
    # .SYNOPSIS
    # Tests whether a repository-relative path contains a .git segment.
    #
    # .DESCRIPTION
    # Compares normalized repository-relative path segments against the exact
    # Git metadata directory segment. Substring matches such as git-helper are
    # intentionally not excluded.
    #
    # .PARAMETER RepositoryRelativePath
    # The repository-relative path to inspect.
    #
    # .EXAMPLE
    # Test-PSScriptAnalyzerGitDirectorySegment -RepositoryRelativePath '.git/hooks/pre-commit.ps1'
    #
    # # Returns true.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [bool] True when an exact .git path segment is present.
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
    [OutputType([bool])]
    param(
        [AllowNull()]
        [object]$RepositoryRelativePath
    )

    Set-StrictMode -Version Latest

    return Test-PSScriptAnalyzerExactPathSegment `
        -RepositoryRelativePath $RepositoryRelativePath `
        -SegmentName '.git'
}

function Test-PSScriptAnalyzerSettingsFilePath {
    # .SYNOPSIS
    # Tests whether a repository-relative path is the analyzer settings file.
    #
    # .DESCRIPTION
    # Normalizes separators to forward slashes and compares against the retained
    # analyzer settings file by repository-relative path. Same-basename files in
    # other directories are intentionally not excluded.
    #
    # .PARAMETER RepositoryRelativePath
    # The repository-relative path to inspect.
    #
    # .EXAMPLE
    # Test-PSScriptAnalyzerSettingsFilePath -RepositoryRelativePath '.github/linting/PSScriptAnalyzerSettings.psd1'
    #
    # # Returns true.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [bool] True when the path is the retained analyzer settings file.
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
    [OutputType([bool])]
    param(
        [AllowNull()]
        [object]$RepositoryRelativePath
    )

    Set-StrictMode -Version Latest

    if ($null -eq $RepositoryRelativePath) {
        return $false
    }

    $strNormalizedPath = ([string]$RepositoryRelativePath).Replace('\', '/')
    return [string]::Equals(
        $strNormalizedPath,
        '.github/linting/PSScriptAnalyzerSettings.psd1',
        [System.StringComparison]::OrdinalIgnoreCase
    )
}

function Get-PSScriptAnalyzerPolicyExclusionReasonCode {
    # .SYNOPSIS
    # Gets the policy-exclusion reason for an analyzer candidate path.
    #
    # .DESCRIPTION
    # Applies repository-relative analyzer input exclusion policy shared by CI
    # discovery, direct candidate classification, and resolved reparse-target
    # validation.
    #
    # .PARAMETER RepositoryRelativePath
    # The repository-relative path to classify.
    #
    # .EXAMPLE
    # Get-PSScriptAnalyzerPolicyExclusionReasonCode -RepositoryRelativePath '.github/linting/PSScriptAnalyzerSettings.psd1'
    #
    # # Returns AnalyzerSettingsFile.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The policy-exclusion reason code, or an empty string when the path
    # is not policy-excluded.
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
    param(
        [AllowNull()]
        [object]$RepositoryRelativePath
    )

    Set-StrictMode -Version Latest

    if (Test-PSScriptAnalyzerSettingsFilePath -RepositoryRelativePath $RepositoryRelativePath) {
        return 'AnalyzerSettingsFile'
    }

    if (Test-PSScriptAnalyzerNodeModuleSegment -RepositoryRelativePath $RepositoryRelativePath) {
        return 'NodeModulesSegment'
    }

    if (Test-PSScriptAnalyzerGitDirectorySegment -RepositoryRelativePath $RepositoryRelativePath) {
        return 'GitDirectorySegment'
    }

    return ''
}

function Get-PSScriptAnalyzerPathComparison {
    # .SYNOPSIS
    # Gets the path comparison mode for the current platform.
    #
    # .DESCRIPTION
    # Returns a conservative string-comparison mode for repository containment
    # checks. Windows paths are compared case-insensitively; other platforms use
    # ordinal comparisons so containment does not broaden unexpectedly.
    #
    # .EXAMPLE
    # Get-PSScriptAnalyzerPathComparison
    #
    # # Returns Ordinal or OrdinalIgnoreCase.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [System.StringComparison] The comparison mode for filesystem paths.
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
    [OutputType([System.StringComparison])]
    param()

    Set-StrictMode -Version Latest

    if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
        return [System.StringComparison]::OrdinalIgnoreCase
    }

    return [System.StringComparison]::Ordinal
}

function Resolve-PSScriptAnalyzerFullPath {
    # .SYNOPSIS
    # Resolves a candidate path to an absolute filesystem path.
    #
    # .DESCRIPTION
    # Combines repository-relative candidate paths with the repository root and
    # canonicalizes the result with GetFullPath without resolving a leaf
    # reparse point.
    #
    # .PARAMETER RepositoryRoot
    # The repository root used to anchor relative candidates.
    #
    # .PARAMETER Path
    # The candidate path to resolve.
    #
    # .EXAMPLE
    # Resolve-PSScriptAnalyzerFullPath -RepositoryRoot '/repo' -Path 'src/a.ps1'
    #
    # # Returns an absolute path under /repo.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The canonical absolute path.
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
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    Set-StrictMode -Version Latest

    $strRepositoryRoot = [System.IO.Path]::GetFullPath($RepositoryRoot)
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath(
        [System.IO.Path]::Combine($strRepositoryRoot, $Path)
    )
}

function Test-PSScriptAnalyzerPathInsideRepository {
    # .SYNOPSIS
    # Tests whether an absolute path is contained by the repository root.
    #
    # .DESCRIPTION
    # Uses Path.GetRelativePath when available and falls back to a
    # canonicalized GetFullPath prefix check that requires a segment boundary.
    #
    # .PARAMETER RepositoryRoot
    # The repository root.
    #
    # .PARAMETER Path
    # The path to validate.
    #
    # .EXAMPLE
    # Test-PSScriptAnalyzerPathInsideRepository -RepositoryRoot '/repo' -Path '/repo/src/a.ps1'
    #
    # # Returns true.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [bool] True when the path is inside the repository root.
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
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    Set-StrictMode -Version Latest

    $strRepositoryRoot = [System.IO.Path]::GetFullPath($RepositoryRoot)
    $strPath = [System.IO.Path]::GetFullPath($Path)

    $objGetRelativePathMethod = [System.IO.Path].GetMethod(
        'GetRelativePath',
        [type[]]@([string], [string])
    )
    if ($null -ne $objGetRelativePathMethod) {
        $strRelativePath = [string]$objGetRelativePathMethod.Invoke(
            $null,
            [object[]]@($strRepositoryRoot, $strPath)
        )
        if ([System.IO.Path]::IsPathRooted($strRelativePath)) {
            return $false
        }

        $arrPart = $strRelativePath.Replace('\', '/').Split([char[]]@('/'))
        return -not ($arrPart | Where-Object { $_ -eq '..' } | Select-Object -First 1)
    }

    $arrTrimCharacter = [char[]]@(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    $strTrimmedRoot = $strRepositoryRoot.TrimEnd($arrTrimCharacter)
    $strTrimmedPath = $strPath.TrimEnd($arrTrimCharacter)
    $objComparison = Get-PSScriptAnalyzerPathComparison

    if ([string]::Equals($strTrimmedRoot, $strTrimmedPath, $objComparison)) {
        return $true
    }

    $strRootPrefix = $strTrimmedRoot + [System.IO.Path]::DirectorySeparatorChar
    return $strTrimmedPath.StartsWith($strRootPrefix, $objComparison)
}

function ConvertTo-PSScriptAnalyzerRepositoryRelativePath {
    # .SYNOPSIS
    # Converts an absolute path to a repository-relative path.
    #
    # .DESCRIPTION
    # Uses Path.GetRelativePath when available and falls back to a
    # canonicalized segment-safe prefix removal. Returned separators are
    # normalized to forward slashes for stable repository-relative rendering.
    #
    # .PARAMETER RepositoryRoot
    # The repository root.
    #
    # .PARAMETER Path
    # The absolute path to convert.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerRepositoryRelativePath -RepositoryRoot '/repo' -Path '/repo/src/a.ps1'
    #
    # # Returns src/a.ps1.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] The repository-relative path.
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
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    Set-StrictMode -Version Latest

    $strRepositoryRoot = [System.IO.Path]::GetFullPath($RepositoryRoot)
    $strPath = [System.IO.Path]::GetFullPath($Path)

    $objGetRelativePathMethod = [System.IO.Path].GetMethod(
        'GetRelativePath',
        [type[]]@([string], [string])
    )
    if ($null -ne $objGetRelativePathMethod) {
        $strRelativePath = [string]$objGetRelativePathMethod.Invoke(
            $null,
            [object[]]@($strRepositoryRoot, $strPath)
        )
        return $strRelativePath.Replace('\', '/')
    }

    if (-not (Test-PSScriptAnalyzerPathInsideRepository -RepositoryRoot $strRepositoryRoot -Path $strPath)) {
        return $strPath
    }

    $arrTrimCharacter = [char[]]@(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    $strTrimmedRoot = $strRepositoryRoot.TrimEnd($arrTrimCharacter)
    $strRelativePath = $strPath.Substring($strTrimmedRoot.Length).TrimStart($arrTrimCharacter)

    return $strRelativePath.Replace('\', '/')
}

function ConvertTo-PSScriptAnalyzerCandidateRecord {
    # .SYNOPSIS
    # Creates a structured analyzer candidate outcome record.
    #
    # .DESCRIPTION
    # Builds the common candidate record shape used by CI and first-adoption
    # reports. EscapedAnalyzerPath is included only for selected candidates, and
    # ResolvedTargetFullName is included only when a reparse target applies.
    #
    # .PARAMETER CandidateFullName
    # The full candidate path.
    #
    # .PARAMETER RepositoryRelativePath
    # The repository-relative candidate path.
    #
    # .PARAMETER OutcomeCategory
    # The outcome category: selected, policy-excluded, or unsafe.
    #
    # .PARAMETER ReasonCode
    # The stable reason code for the outcome.
    #
    # .PARAMETER EscapedAnalyzerPath
    # The wildcard-escaped candidate path for selected analyzer input.
    #
    # .PARAMETER ResolvedTargetFullName
    # The resolved reparse target path, when applicable.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerCandidateRecord -CandidateFullName '/repo/a.ps1' -RepositoryRelativePath 'a.ps1' -OutcomeCategory selected -ReasonCode Selected
    #
    # # Returns one selected candidate record.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [pscustomobject] A candidate outcome record.
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
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [string]$CandidateFullName,

        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [string]$RepositoryRelativePath,

        [Parameter(Mandatory = $true)]
        [ValidateSet('selected', 'policy-excluded', 'unsafe')]
        [string]$OutcomeCategory,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ReasonCode,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$EscapedAnalyzerPath,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$ResolvedTargetFullName
    )

    Set-StrictMode -Version Latest

    $hashtableRecord = [ordered]@{
        CandidateFullName = $CandidateFullName
        RepositoryRelativePath = $RepositoryRelativePath
        OutcomeCategory = $OutcomeCategory
        ReasonCode = $ReasonCode
    }

    if ($OutcomeCategory -eq 'selected') {
        $hashtableRecord['EscapedAnalyzerPath'] = $EscapedAnalyzerPath
    }

    if (-not [string]::IsNullOrEmpty($ResolvedTargetFullName)) {
        $hashtableRecord['ResolvedTargetFullName'] = $ResolvedTargetFullName
    }

    return [pscustomobject]$hashtableRecord
}

function Resolve-PSScriptAnalyzerCandidate {
    # .SYNOPSIS
    # Classifies one PSScriptAnalyzer candidate path.
    #
    # .DESCRIPTION
    # Applies shared analyzer selection policy to one candidate, including exact
    # node_modules and .git segment exclusions, analyzer settings-file
    # exclusion, allowed-extension validation, repository containment checks,
    # fail-closed leaf reparse-point handling, and wildcard escaping for
    # selected Invoke-ScriptAnalyzer input.
    #
    # .PARAMETER RepositoryRoot
    # The repository root used for containment and relative path rendering.
    #
    # .PARAMETER CandidatePath
    # The candidate path. Relative values are anchored to RepositoryRoot.
    #
    # .PARAMETER RepositoryRelativePath
    # Optional repository-relative path already discovered lexically, such as a
    # raw Git path. When supplied, this exact decoded string is preserved in the
    # structured record.
    #
    # .EXAMPLE
    # Resolve-PSScriptAnalyzerCandidate -RepositoryRoot '/repo' -CandidatePath 'src/a.ps1'
    #
    # # Returns a selected, policy-excluded, or unsafe record.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [pscustomobject] A candidate outcome record.
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
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CandidatePath,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$RepositoryRelativePath
    )

    Set-StrictMode -Version Latest

    $strRepositoryRoot = [System.IO.Path]::GetFullPath($RepositoryRoot)
    $strCandidateFullName = Resolve-PSScriptAnalyzerFullPath `
        -RepositoryRoot $strRepositoryRoot `
        -Path $CandidatePath

    $strRepositoryRelativePath = $RepositoryRelativePath
    if ([string]::IsNullOrEmpty($strRepositoryRelativePath)) {
        $strRepositoryRelativePath = ConvertTo-PSScriptAnalyzerRepositoryRelativePath `
            -RepositoryRoot $strRepositoryRoot `
            -Path $strCandidateFullName
    }

    $strPolicyExclusionReasonCode = Get-PSScriptAnalyzerPolicyExclusionReasonCode `
        -RepositoryRelativePath $strRepositoryRelativePath
    if (-not [string]::IsNullOrEmpty($strPolicyExclusionReasonCode)) {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'policy-excluded' `
            -ReasonCode $strPolicyExclusionReasonCode
    }

    if (-not (Test-PSScriptAnalyzerCandidateExtension -Path $strRepositoryRelativePath)) {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'policy-excluded' `
            -ReasonCode 'ExtensionNotAllowed'
    }

    if (-not (Test-PSScriptAnalyzerPathInsideRepository -RepositoryRoot $strRepositoryRoot -Path $strCandidateFullName)) {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode 'CandidateOutsideRepository'
    }

    try {
        $objCandidateItem = Get-Item -LiteralPath $strCandidateFullName -Force -ErrorAction Stop
    } catch {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode 'MissingCandidate'
    }

    if ($objCandidateItem.PSIsContainer) {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode 'CandidateIsDirectory'
    }

    $boolCandidateIsReparsePoint = (
        ($objCandidateItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
    )

    if (-not $boolCandidateIsReparsePoint) {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'selected' `
            -ReasonCode 'Selected' `
            -EscapedAnalyzerPath ([System.Management.Automation.WildcardPattern]::Escape($strCandidateFullName))
    }

    $objResolveLinkTargetMethod = $objCandidateItem.GetType().GetMethod(
        'ResolveLinkTarget',
        [type[]]@([bool])
    )
    if ($null -eq $objResolveLinkTargetMethod) {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode 'ReparsePointResolutionUnsupported'
    }

    try {
        $objResolvedTarget = $objResolveLinkTargetMethod.Invoke(
            $objCandidateItem,
            [object[]]@($true)
        )
    } catch {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode 'ReparsePointResolutionFailed'
    }

    if ($null -eq $objResolvedTarget) {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode 'ReparsePointResolutionFailed'
    }

    $strResolvedTargetFullName = [System.IO.Path]::GetFullPath($objResolvedTarget.FullName)
    try {
        $objResolvedTargetItem = Get-Item -LiteralPath $strResolvedTargetFullName -Force -ErrorAction Stop
    } catch {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode 'MissingTarget' `
            -ResolvedTargetFullName $strResolvedTargetFullName
    }

    if (-not (Test-PSScriptAnalyzerPathInsideRepository -RepositoryRoot $strRepositoryRoot -Path $strResolvedTargetFullName)) {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode 'TargetOutsideRepository' `
            -ResolvedTargetFullName $strResolvedTargetFullName
    }

    if ($objResolvedTargetItem.PSIsContainer) {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode 'TargetIsDirectory' `
            -ResolvedTargetFullName $strResolvedTargetFullName
    }

    if (-not (Test-PSScriptAnalyzerCandidateExtension -Path $strResolvedTargetFullName)) {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode 'TargetExtensionNotAllowed' `
            -ResolvedTargetFullName $strResolvedTargetFullName
    }

    $strResolvedTargetRelativePath = ConvertTo-PSScriptAnalyzerRepositoryRelativePath `
        -RepositoryRoot $strRepositoryRoot `
        -Path $strResolvedTargetFullName

    $strTargetPolicyExclusionReasonCode = Get-PSScriptAnalyzerPolicyExclusionReasonCode `
        -RepositoryRelativePath $strResolvedTargetRelativePath
    if (-not [string]::IsNullOrEmpty($strTargetPolicyExclusionReasonCode)) {
        return ConvertTo-PSScriptAnalyzerCandidateRecord `
            -CandidateFullName $strCandidateFullName `
            -RepositoryRelativePath $strRepositoryRelativePath `
            -OutcomeCategory 'unsafe' `
            -ReasonCode ('Target{0}' -f $strTargetPolicyExclusionReasonCode) `
            -ResolvedTargetFullName $strResolvedTargetFullName
    }

    return ConvertTo-PSScriptAnalyzerCandidateRecord `
        -CandidateFullName $strCandidateFullName `
        -RepositoryRelativePath $strRepositoryRelativePath `
        -OutcomeCategory 'selected' `
        -ReasonCode 'SelectedReparsePoint' `
        -EscapedAnalyzerPath ([System.Management.Automation.WildcardPattern]::Escape($strCandidateFullName)) `
        -ResolvedTargetFullName $strResolvedTargetFullName
}

function Get-PSScriptAnalyzerCandidate {
    # .SYNOPSIS
    # Finds and classifies PSScriptAnalyzer candidates under a repository root.
    #
    # .DESCRIPTION
    # Performs the CI directory-walk layer for analyzer candidates. The walk can
    # include hidden directories with Force, prunes exact node_modules and .git
    # directory segments before descent, does not traverse directory reparse
    # points, and delegates each selected leaf to the shared per-path
    # classifier.
    #
    # .PARAMETER RepositoryRoot
    # The repository root to scan.
    #
    # .PARAMETER DirectoryVisibility
    # VisibleOnly enumerates without Force. All enumerates with Force so dot
    # directories and attribute-hidden directories are visible before policy
    # pruning.
    #
    # .EXAMPLE
    # Get-PSScriptAnalyzerCandidate -RepositoryRoot '/repo'
    #
    # # Streams candidate classification records.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [pscustomobject] Candidate classification records.
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
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RepositoryRoot,

        [ValidateSet('VisibleOnly', 'All')]
        [string]$DirectoryVisibility = 'VisibleOnly'
    )

    Set-StrictMode -Version Latest

    $strRepositoryRoot = [System.IO.Path]::GetFullPath($RepositoryRoot)
    $queueDirectory = [System.Collections.Generic.Queue[string]]::new()
    $queueDirectory.Enqueue($strRepositoryRoot)

    while ($queueDirectory.Count -gt 0) {
        $strDirectoryPath = $queueDirectory.Dequeue()

        $hashtableDirectoryParameter = @{
            LiteralPath = $strDirectoryPath
            Directory = $true
        }
        $hashtableFileParameter = @{
            LiteralPath = $strDirectoryPath
            File = $true
        }
        if ($DirectoryVisibility -eq 'All') {
            $hashtableDirectoryParameter['Force'] = $true
            $hashtableFileParameter['Force'] = $true
        }

        foreach ($objDirectory in @(Get-ChildItem @hashtableDirectoryParameter)) {
            $strDirectoryRelativePath = ConvertTo-PSScriptAnalyzerRepositoryRelativePath `
                -RepositoryRoot $strRepositoryRoot `
                -Path $objDirectory.FullName
            $strDirectoryPolicyExclusionReasonCode = Get-PSScriptAnalyzerPolicyExclusionReasonCode `
                -RepositoryRelativePath $strDirectoryRelativePath
            if (
                ($strDirectoryPolicyExclusionReasonCode -eq 'NodeModulesSegment') -or
                ($strDirectoryPolicyExclusionReasonCode -eq 'GitDirectorySegment')
            ) {
                continue
            }

            $boolDirectoryIsReparsePoint = (
                ($objDirectory.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
            )
            if ($boolDirectoryIsReparsePoint) {
                continue
            }

            $queueDirectory.Enqueue($objDirectory.FullName)
        }

        foreach ($objFile in @(Get-ChildItem @hashtableFileParameter)) {
            $strFileRelativePath = ConvertTo-PSScriptAnalyzerRepositoryRelativePath `
                -RepositoryRoot $strRepositoryRoot `
                -Path $objFile.FullName
            if (-not (Test-PSScriptAnalyzerCandidateExtension -Path $strFileRelativePath)) {
                continue
            }

            Resolve-PSScriptAnalyzerCandidate `
                -RepositoryRoot $strRepositoryRoot `
                -CandidatePath $objFile.FullName `
                -RepositoryRelativePath $strFileRelativePath
        }
    }
}

function Get-PSScriptAnalyzerCandidateSummary {
    # .SYNOPSIS
    # Groups candidate records by outcome category.
    #
    # .DESCRIPTION
    # Returns explicit child arrays and summary counts for selected,
    # policy-excluded, and unsafe candidates so JSON serialization preserves the
    # documented empty, singleton, and multi-item array shapes.
    #
    # .PARAMETER Candidate
    # Candidate classification records to summarize.
    #
    # .EXAMPLE
    # Get-PSScriptAnalyzerCandidateSummary -Candidate $candidateRecords
    #
    # # Returns grouped candidate data and counts.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [pscustomobject] Candidate groups and counts.
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
    [OutputType([pscustomobject])]
    param(
        [AllowNull()]
        [object[]]$Candidate
    )

    Set-StrictMode -Version Latest

    $listSelected = [System.Collections.Generic.List[pscustomobject]]::new()
    $listPolicyExcluded = [System.Collections.Generic.List[pscustomobject]]::new()
    $listUnsafe = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($objCandidate in @($Candidate)) {
        if ($null -eq $objCandidate) {
            continue
        }

        if ($objCandidate.OutcomeCategory -eq 'selected') {
            [void]($listSelected.Add($objCandidate))
        } elseif ($objCandidate.OutcomeCategory -eq 'policy-excluded') {
            [void]($listPolicyExcluded.Add($objCandidate))
        } elseif ($objCandidate.OutcomeCategory -eq 'unsafe') {
            [void]($listUnsafe.Add($objCandidate))
        } else {
            throw ("Unexpected PSScriptAnalyzer candidate OutcomeCategory: {0}" -f $objCandidate.OutcomeCategory)
        }
    }

    return [pscustomobject]@{
        Selected = [object[]]$listSelected.ToArray()
        PolicyExcluded = [object[]]$listPolicyExcluded.ToArray()
        Unsafe = [object[]]$listUnsafe.ToArray()
        SummaryCounts = [pscustomobject]@{
            Selected = $listSelected.Count
            PolicyExcluded = $listPolicyExcluded.Count
            Unsafe = $listUnsafe.Count
        }
    }
}

function ConvertTo-PSScriptAnalyzerCandidateSummaryLine {
    # .SYNOPSIS
    # Renders candidate summary lines.
    #
    # .DESCRIPTION
    # Converts structured candidate outcome data into safe one-line text for CI
    # logs and first-adoption reports. Structured candidate fields remain
    # unchanged; only human-facing rendering is escaped.
    #
    # .PARAMETER CandidateSummary
    # The grouped candidate summary object.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerCandidateSummaryLine -CandidateSummary $summary
    #
    # # Streams one or more summary lines.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] One-line human-readable candidate summaries.
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
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$CandidateSummary
    )

    Set-StrictMode -Version Latest

    $intSelectedCount = [int]$CandidateSummary.SummaryCounts.Selected
    $intPolicyExcludedCount = [int]$CandidateSummary.SummaryCounts.PolicyExcluded
    $intUnsafeCount = [int]$CandidateSummary.SummaryCounts.Unsafe

    $arrSummaryArgument = @(
        $intSelectedCount
        $intPolicyExcludedCount
        $intUnsafeCount
    )
    'PSScriptAnalyzer candidates: {0} selected; {1} policy-excluded; {2} unsafe.' -f $arrSummaryArgument

    if (($intSelectedCount -eq 0) -and ($intPolicyExcludedCount -gt 0) -and ($intUnsafeCount -eq 0)) {
        'No analyzer inputs were selected; {0} candidate path(s) were excluded by policy.' -f `
            $intPolicyExcludedCount
    }

    if ($intPolicyExcludedCount -gt 0) {
        $arrReasonCount = @(
            $CandidateSummary.PolicyExcluded |
                Group-Object -Property ReasonCode |
                Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, Name |
                ForEach-Object { '{0} ({1})' -f $_.Name, $_.Count }
        )
        'Policy-excluded reason counts: {0}.' -f ($arrReasonCount -join '; ')

        $arrExample = @(
            $CandidateSummary.PolicyExcluded |
                Select-Object -First 3 |
                ForEach-Object {
                    ConvertTo-PSScriptAnalyzerCandidateSingleLineText -Value $_.RepositoryRelativePath
                }
        )
        $arrPolicyExcludedArgument = @(
            $intPolicyExcludedCount
            ($arrExample -join '; ')
        )
        'Policy-excluded candidates: {0} (examples: {1}).' -f $arrPolicyExcludedArgument
    }

    foreach ($objUnsafeCandidate in @($CandidateSummary.Unsafe)) {
        $strCandidatePath = ConvertTo-PSScriptAnalyzerCandidateSingleLineText `
            -Value $objUnsafeCandidate.RepositoryRelativePath
        $strReasonCode = ConvertTo-PSScriptAnalyzerCandidateSingleLineText `
            -Value $objUnsafeCandidate.ReasonCode
        $strLine = 'Unsafe candidate: {0}; reason: {1}' -f $strCandidatePath, $strReasonCode
        if ($objUnsafeCandidate.PSObject.Properties.Name -contains 'ResolvedTargetFullName') {
            $strResolvedTarget = ConvertTo-PSScriptAnalyzerCandidateSingleLineText `
                -Value $objUnsafeCandidate.ResolvedTargetFullName
            $strLine = '{0}; resolved target: {1}' -f $strLine, $strResolvedTarget
        }
        $strLine
    }
}

function ConvertTo-PSScriptAnalyzerCandidateGitHubErrorCommand {
    # .SYNOPSIS
    # Renders an unsafe candidate as a GitHub Actions error command.
    #
    # .DESCRIPTION
    # Converts one unsafe candidate record into a GitHub Actions annotation
    # command with command-message escaping for all dynamic candidate fields.
    #
    # .PARAMETER Candidate
    # The unsafe candidate record to render.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerCandidateGitHubErrorCommand -Candidate $candidate
    #
    # # Returns a GitHub Actions error command.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] A GitHub Actions error command.
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
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$Candidate
    )

    Set-StrictMode -Version Latest

    $arrMessageArgument = @(
        $Candidate.RepositoryRelativePath
        $Candidate.ReasonCode
    )
    $strMessage = 'Unsafe PSScriptAnalyzer candidate: {0}; reason: {1}' -f $arrMessageArgument
    if ($Candidate.PSObject.Properties.Name -contains 'ResolvedTargetFullName') {
        $strMessage = '{0}; resolved target: {1}' -f $strMessage, $Candidate.ResolvedTargetFullName
    }

    return '::error title=Unsafe PSScriptAnalyzer candidate::{0}' -f (
        ConvertTo-PSScriptAnalyzerCandidateGitHubCommandValue -Value $strMessage
    )
}

function ConvertTo-PSScriptAnalyzerCandidateAzurePipelinesErrorCommand {
    # .SYNOPSIS
    # Renders an unsafe candidate as an Azure Pipelines error command.
    #
    # .DESCRIPTION
    # Converts one unsafe candidate record into an Azure Pipelines logging
    # command with property and message escaping for dynamic candidate fields.
    #
    # .PARAMETER Candidate
    # The unsafe candidate record to render.
    #
    # .EXAMPLE
    # ConvertTo-PSScriptAnalyzerCandidateAzurePipelinesErrorCommand -Candidate $candidate
    #
    # # Returns an Azure Pipelines error command.
    #
    # .INPUTS
    # None. This function does not accept pipeline input.
    #
    # .OUTPUTS
    # [string] An Azure Pipelines logging command.
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
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [pscustomobject]$Candidate
    )

    Set-StrictMode -Version Latest

    $arrMessageArgument = @(
        $Candidate.RepositoryRelativePath
        $Candidate.ReasonCode
    )
    $strMessage = 'Unsafe PSScriptAnalyzer candidate: {0}; reason: {1}' -f $arrMessageArgument
    if ($Candidate.PSObject.Properties.Name -contains 'ResolvedTargetFullName') {
        $strMessage = '{0}; resolved target: {1}' -f $strMessage, $Candidate.ResolvedTargetFullName
    }

    $arrCommandArgument = @(
        (ConvertTo-PSScriptAnalyzerCandidateAzureCommandPropertyValue -Value $Candidate.CandidateFullName)
        (ConvertTo-PSScriptAnalyzerCandidateAzureCommandPropertyValue -Value $Candidate.ReasonCode)
        (ConvertTo-PSScriptAnalyzerCandidateAzureCommandMessage -Value $strMessage)
    )
    return '##vso[task.logissue type=error;sourcepath={0};code={1};]{2}' -f $arrCommandArgument
}
