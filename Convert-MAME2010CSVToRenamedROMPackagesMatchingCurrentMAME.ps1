# Convert-MAME2010CSVToRenamedROMPackagesMatchingCurrentMAME.ps1
# MAME 2010 is built from ROM set version 0.139

$strThisScriptVersionNumber = [version]'1.0.20201012.0'

#region License
###############################################################################################
# Copyright 2020 Frank Lesniak

# Permission is hereby granted, free of charge, to any person obtaining a copy of this software
# and associated documentation files (the "Software"), to deal in the Software without
# restriction, including without limitation the rights to use, copy, modify, merge, publish,
# distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the
# Software is furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all copies or
# substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
# BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
# DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
###############################################################################################
#endregion License

#region DownloadLocationNotice
# The most up-to-date version of this script can be found on the author's GitHub repository
# at https://github.com/franklesniak/ROMSorter
#endregion DownloadLocationNotice

$actionPreferenceNewVerbose = $VerbosePreference
$actionPreferenceFormerVerbose = $VerbosePreference
$actionPreferenceNewDebug = $DebugPreference
$actionPreferenceFormerDebug = $DebugPreference
$strPrerequisiteDownloadPageURL = $null
$strPrerequisiteCSVURL = $null
$strLocalPrerequisiteCSVFilePath = $null
$strExclusionDownloadPageURL = $null
$strExclusionCSVURL = $null
$strLocalExclusionCSVFilePath = $null
$strSubfolderPath = $null

#region Inputs
###############################################################################################
# This script requires the MAME 2010 DAT converted to CSV. Set the following path to point
# to the corresponding CSV:
$strPathToROMPackageMetadataCSV = Join-Path '.' 'MAME_2010_DAT.csv'

# Set the following to be the actual version number of the MAME ROM set
$strSourceMAMEVersion = '0.139'

# This script also requires the Progetto Snaps RenameSet converted to a CSV. Set the following
# path to point to the corresponding CSV:
$strPathToProgettoSnapsRenameSetInCSVFormat = Join-Path '.' 'Progetto_Snaps_RenameSet.csv'

# If it is necessary to add prerequisite ROM package deletions or renames to the RenameSet for
# this emulator version/ROM set (i.e., delete or rename ROM packages before applying the
# RenameSet), then set the following URLs to point to the download location (or, if you do not
# wish to perform prerequisite deletions/renames, comment these lines out):
$strPrerequisiteDownloadPageURL = 'https://github.com/franklesniak/ROMSorter'
$strPrerequisiteCSVURL = 'https://raw.githubusercontent.com/franklesniak/ROMSorter/master/MAME_2010_RenameSet_Prerequisites.csv'
$strSubfolderPath = Join-Path '.' 'MAME_2010_Resources'
# If you want the script to use a local copy of the prerequisite file instead of downloading it
# from the URL above, uncomment and configure the following line:
# $strLocalPrerequisiteCSVFilePath = Join-Path $strSubfolderPath 'MAME_2010_RenameSet_Prerequisites.csv'

# If it is necessary to exclude items in the RenameSet for this emulator version/ROM set, then
# set the following URLs to point to the download location (if you do not wish to exclude
# anything from the RenameSet, comment these lines out):
$strExclusionDownloadPageURL = 'https://github.com/franklesniak/ROMSorter'
$strExclusionCSVURL = 'https://raw.githubusercontent.com/franklesniak/ROMSorter/master/MAME_2010_RenameSet_Exclusions.csv'
$strSubfolderPath = Join-Path '.' 'MAME_2010_Resources'
# If you want the script to use a local copy of the exclusions file instead of downloading it
# from the URL above, uncomment and configure the following line:
# $strLocalExclusionCSVFilePath = Join-Path $strSubfolderPath 'MAME_2010_RenameSet_Exclusions.csv'

# The file will be processed and output as a CSV to
# .\Progetto_Snaps_RenameSet.csv
# or if on Linux / MacOS: ./Progetto_Snaps_RenameSet.csv
$strCSVOutputFile = Join-Path '.' 'MAME_2010_DAT_With_Time-Advanced_ROM_Package_Names.csv'

# Comment-out the following line if you prefer that the script operate silently.
$actionPreferenceNewVerbose = [System.Management.Automation.ActionPreference]::Continue

# Remove the comment from the following line if you prefer that the script output extra
# debugging information.
# $actionPreferenceNewDebug = [System.Management.Automation.ActionPreference]::Continue


###############################################################################################
#endregion Inputs

function Split-StringOnLiteralString {
    # This function takes two positional arguments
    # The first argument is a string, and the string to be split
    # The second argument is a string or char, and it is that which is to split the string in the first parameter
    #
    # Note: This function always returns an array, even when there is zero or one element in it.
    #
    # Example:
    # $result = Split-StringOnLiteralString 'foo' ' '
    # # $result.GetType().FullName is System.Object[]
    # # $result.Count is 1
    #
    # Example 2:
    # $result = Split-StringOnLiteralString 'What do you think of this function?' ' '
    # # $result.Count is 7

    $strThisFunctionVersionNumber = [version]'2.0.20200820.0'

    trap {
        Write-Error 'An error occurred using the Split-StringOnLiteralString function. This was most likely caused by the arguments supplied not being strings'
    }

    if ($args.Length -ne 2) {
        Write-Error 'Split-StringOnLiteralString was called without supplying two arguments. The first argument should be the string to be split, and the second should be the string or character on which to split the string.'
        $result = @()
    } else {
        $objToSplit = $args[0]
        $objSplitter = $args[1]
        if ($null -eq $objToSplit) {
            $result = @()
        } elseif ($null -eq $objSplitter) {
            # Splitter was $null; return string to be split within an array (of one element).
            $result = @($objToSplit)
        } else {
            if ($objToSplit.GetType().Name -ne 'String') {
                Write-Warning 'The first argument supplied to Split-StringOnLiteralString was not a string. It will be attempted to be converted to a string. To avoid this warning, cast arguments to a string before calling Split-StringOnLiteralString.'
                $strToSplit = [string]$objToSplit
            } else {
                $strToSplit = $objToSplit
            }

            if (($objSplitter.GetType().Name -ne 'String') -and ($objSplitter.GetType().Name -ne 'Char')) {
                Write-Warning 'The second argument supplied to Split-StringOnLiteralString was not a string. It will be attempted to be converted to a string. To avoid this warning, cast arguments to a string before calling Split-StringOnLiteralString.'
                $strSplitter = [string]$objSplitter
            } elseif ($objSplitter.GetType().Name -eq 'Char') {
                $strSplitter = [string]$objSplitter
            } else {
                $strSplitter = $objSplitter
            }

            $strSplitterInRegEx = [regex]::Escape($strSplitter)

            # With the leading comma, force encapsulation into an array so that an array is
            # returned even when there is one element:
            $result = @([regex]::Split($strToSplit, $strSplitterInRegEx))
        }
    }

    # The following code forces the function to return an array, always, even when there are
    # zero or one elements in the array
    $intElementCount = 1
    if ($null -ne $result) {
        if ($result.GetType().FullName.Contains('[]')) {
            if (($result.Count -ge 2) -or ($result.Count -eq 0)) {
                $intElementCount = $result.Count
            }
        }
    }
    $strLowercaseFunctionName = $MyInvocation.InvocationName.ToLower()
    $boolArrayEncapsulation = $MyInvocation.Line.ToLower().Contains('@(' + $strLowercaseFunctionName + ')') -or $MyInvocation.Line.ToLower().Contains('@(' + $strLowercaseFunctionName + ' ')
    if ($boolArrayEncapsulation) {
        $result
    } elseif ($intElementCount -eq 0) {
        , @()
    } elseif ($intElementCount -eq 1) {
        , (, ($args[0]))
    } else {
        $result
    }
}

function New-BackwardCompatibleCaseInsensitiveHashtable {
    # New-BackwardCompatibleCaseInsensitiveHashtable is designed to create a case-insensitive
    # hashtable that is backward-compatible all the way to PowerShell v1, yet forward-
    # compatible to all versions of PowerShell. It replaces other constructors on newer
    # versions of PowerShell such as:
    # $hashtable = @{}
    # This function is useful if you need to work with hashtables (key-value pairs), but also
    # need your code to be able to run on any version of PowerShell.
    #
    # Usage:
    # $hashtable = New-BackwardCompatibleCaseInsensitiveHashtable

    $strThisFunctionVersionNumber = [version]'1.0.20200817.0'

    $cultureDoNotCare = [System.Globalization.CultureInfo]::InvariantCulture
    $caseInsensitiveHashCodeProvider = New-Object -TypeName 'System.Collections.CaseInsensitiveHashCodeProvider' -ArgumentList @($cultureDoNotCare)
    $caseInsensitiveComparer = New-Object -TypeName 'System.Collections.CaseInsensitiveComparer' -ArgumentList @($cultureDoNotCare)
    New-Object -TypeName 'System.Collections.Hashtable' -ArgumentList @($caseInsensitiveHashCodeProvider, $caseInsensitiveComparer)
}

function Convert-MAMEVersionNumberToRepresentativePowerShellVersion {
    $strMAMEVersion = $args[0]
    $strMAMEVersion = $strMAMEVersion.ToLower()

    switch ($strMAMEVersion) {
        '0.26a' { $strScrubbedMAMEVersion = '0.26u1' }
        '0.35fix' { $strScrubbedMAMEVersion = '0.35.1' }
        '0.69a' { $strScrubbedMAMEVersion = '0.69u1' }
        '0.69b' { $strScrubbedMAMEVersion = '0.69u2' }
        '0.71u3p' { $strScrubbedMAMEVersion = '0.71u3' }
        '0.124a' { $strScrubbedMAMEVersion = '0.124.1' }
        default { $strScrubbedMAMEVersion = $strMAMEVersion }
    }

    $boolBetaVersion = $strScrubbedMAMEVersion.Contains('b')
    $boolReleaseCandidateVersion = $strScrubbedMAMEVersion.Contains('rc')
    $boolUpdateVersion = $strScrubbedMAMEVersion.Contains('u')
    $arrVersionSplit = Split-StringOnLiteralString $strScrubbedMAMEVersion '.'
    $boolPointReleaseVersion = ($arrVersionSplit.Count -ge 3)

    if ($boolBetaVersion) {
        $arrVersionSplit = Split-StringOnLiteralString $strScrubbedMAMEVersion 'b'
        $strFirstAndSecondPieceOfVersion = $arrVersionSplit[0]
        $strThirdPieceOfVersion = [string]([int]($arrVersionSplit[$arrVersionSplit.Length - 1]) + 0)
        $strFourthPieceOfVersion = '0'
    } elseif ($boolReleaseCandidateVersion) {
        $arrVersionSplit = Split-StringOnLiteralString $strScrubbedMAMEVersion 'rc'
        $strFirstAndSecondPieceOfVersion = $arrVersionSplit[0]
        $strThirdPieceOfVersion = [string]([int]($arrVersionSplit[$arrVersionSplit.Length - 1]) + 50)
        $strFourthPieceOfVersion = '0'
    } elseif ($boolUpdateVersion) {
        $arrVersionSplit = Split-StringOnLiteralString $strScrubbedMAMEVersion 'u'
        $strFirstAndSecondPieceOfVersion = $arrVersionSplit[0]
        $strThirdPieceOfVersion = [string]([int]($arrVersionSplit[$arrVersionSplit.Length - 1]) + 100)
        $strFourthPieceOfVersion = '0'
    } elseif ($boolPointReleaseVersion) {
        $arrVersionSplit = Split-StringOnLiteralString $strScrubbedMAMEVersion '.'
        $strFirstAndSecondPieceOfVersion = ""
        for ($intCounterA = 0; $intCounterA -le ($arrVersionSplit.Count - 2); $intCounterA++) {
            $strFirstAndSecondPieceOfVersion += $arrVersionSplit[$intCounterA] + '.'
        }
        $strFirstAndSecondPieceOfVersion = $strFirstAndSecondPieceOfVersion.Substring(0, $strFirstAndSecondPieceOfVersion.Length - 1)
        $strThirdPieceOfVersion = '100'
        $strFourthPieceOfVersion = $arrVersionSplit[$arrVersionSplit.Length - 1]
    } else {
        $strFirstAndSecondPieceOfVersion = $strScrubbedMAMEVersion
        $strThirdPieceOfVersion = '100'
        $strFourthPieceOfVersion = '0'
    }

    $strPowerShellFriendlyVersion = $strFirstAndSecondPieceOfVersion + '.' + $strThirdPieceOfVersion + '.' + $strFourthPieceOfVersion
    [version]$strPowerShellFriendlyVersion
}

function Get-AbsoluteURLFromRelative {
    # This functions takes a potentially relative URL (/etc/foo.html) and turns it into an
    # absolute URL if it is, in fact, a relative URL. If the URL is an absolute URL, then the
    # function simply returns the absolute URL.
    #
    # The function takes two positional arguments.
    #
    # The first argument is a string containing the base URL (i.e., the parent URL) from which
    # the second URL is derived
    #
    # The second argument is a string containing the (potentially) relative URL.
    #
    # Example 1:
    # $strURLBase = 'http://foo.net/stuff/index.html'
    # $strURLRelative = '/downloads/list.txt'
    # $strURLAbsolute = Get-AbsoluteURLFromRelative $strURLBase $strURLRelative
    # # $strURLAbsolute is 'http://foo.net/downloads/list.txt'
    #
    # Example 2:
    # $strURLBase = 'http://foo.net/stuff/index.html'
    # $strURLRelative = 'downloads/list.txt'
    # $strURLAbsolute = Get-AbsoluteURLFromRelative $strURLBase $strURLRelative
    # # $strURLAbsolute is 'http://foo.net/stuff/downloads/list.txt'
    #
    # Example 3:
    # $strURLBase = 'http://foo.net/stuff/index.html'
    # $strURLRelative = 'http://foo.net/stuff/downloads/list.txt'
    # $strURLAbsolute = Get-AbsoluteURLFromRelative $strURLBase $strURLRelative
    # # $strURLAbsolute is 'http://foo.net/stuff/downloads/list.txt'
    #
    # Note: this function is converted from https://stackoverflow.com/a/34603567/2134110
    # Thanks to Vikash Rathee for pointing me in the right direction

    $strURLBase = $args[0]
    $strURLRelative = $args[1]

    $strThisFunctionVersionNumber = [version]'1.0.20201004.0'

    $uriKindRelativeOrAbsolute = [System.UriKind]::RelativeOrAbsolute
    $uriWorking = New-Object -TypeName 'System.Uri' -ArgumentList @($strURLRelative, $uriKindRelativeOrAbsolute)
    if ($uriWorking.IsAbsoluteUri -ne $true) {
        $uriBase = New-Object -TypeName 'System.Uri' -ArgumentList @($strURLBase)
        $uriWorking = New-Object -TypeName 'System.Uri' -ArgumentList @($uriBase, $strURLRelative)
    }
    $uriWorking.ToString()
}

function Test-RenameSetRecordShouldBeIgnored {
    $refArrRenameSetExclusions = $args[0]
    $refPSCustomObjectRenameSetItem = $args[1]

    # $arrRenameSetExclusions = $refArrRenameSetExclusions.Value
    # $PSCustomObjectRenameSetItem = $refPSCustomObjectRenameSetItem.Value

    $PSCustomObjectRenameSetPossibleExclusionMatch = ($refPSCustomObjectRenameSetItem.Value) |
        Select-Object -Property @('MAMEVersionPowerShellFriendly', 'Operation', 'OldROMPackageName', 'NewROMPackageName')

    $arrMatchedExclusions = @(($refArrRenameSetExclusions.Value) |
            Where-Object { $_.MAMEVersionPowerShellFriendly -eq $PSCustomObjectRenameSetPossibleExclusionMatch.MAMEVersionPowerShellFriendly } |
            Where-Object { $_.Operation -eq $PSCustomObjectRenameSetPossibleExclusionMatch.Operation } |
            Where-Object { $_.OldROMPackageName -eq $PSCustomObjectRenameSetPossibleExclusionMatch.OldROMPackageName } |
            Where-Object { $_.NewROMPackageName -eq $PSCustomObjectRenameSetPossibleExclusionMatch.NewROMPackageName })

    ($arrMatchedExclusions.Count -ne 0)
}

$VerbosePreference = $actionPreferenceNewVerbose
$DebugPreference = $actionPreferenceNewDebug

$boolErrorOccurred = $false

if ((Test-Path $strPathToROMPackageMetadataCSV) -ne $true) {
    Write-Error ('The input file "' + $strPathToROMPackageMetadataCSV + '" is missing. Please generate it using the corresponding "Convert-..." script and then re-run this script')
    $boolErrorOccurred = $true
}

if ((Test-Path $strPathToProgettoSnapsRenameSetInCSVFormat) -ne $true) {
    Write-Error ('The input file "' + $strPathToProgettoSnapsRenameSetInCSVFormat + '" is missing. Please generate it using the "Convert-ProgettoSnapsRenameSetIniToCsv.ps1" script and then re-run this script')
    $boolErrorOccurred = $true
}

if ($boolErrorOccurred -eq $true) {
    break
}

# Get the prerequisite CSV
if ($null -eq $strLocalPrerequisiteCSVFilePath -and $null -eq $strPrerequisiteDownloadPageURL -and $null -eq $strPrerequisiteCSVURL) {
    $arrRenameSetPrerequisites = @()
} else {
    $arrCommands = @(Get-Command Invoke-WebRequest)
    $boolInvokeWebRequestAvailable = ($arrCommands.Count -ge 1)
    if ($null -eq $strLocalPrerequisiteCSVFilePath -and $boolInvokeWebRequestAvailable) {
        $VerbosePreference = $actionPreferenceFormerVerbose
        $arrModules = @(Get-Module PowerHTML -ListAvailable)
        $VerbosePreference = $actionPreferenceNewVerbose
        if ($arrModules.Count -eq 0) {
            Write-Warning 'It is recommended that you install the PowerHTML module using "Install-Module PowerHTML" before continuing. Doing so will allow this script to obtain the URL for the most-current RenameSet prerequisite file automatically. Without PowerHTML, this script is using a potentially-outdated URL. Break out of ths script now to install PowerHTML, then re-run the script'
            $strEffectiveURL = $strPrerequisiteCSVURL
        } else {
            Write-Verbose ('Parsing site ' + $strPrerequisiteDownloadPageURL + ' to dynamically obtain RenameSet prerequisite download URL...')
            $arrLoadedModules = @(Get-Module PowerHTML)
            if ($arrLoadedModules.Count -eq 0) {
                $VerbosePreference = $actionPreferenceFormerVerbose
                Import-Module PowerHTML
                $VerbosePreference = $actionPreferenceNewVerbose
            }

            $strNextDownloadPageURL = $strPrerequisiteDownloadPageURL
            $HtmlNodeDownloadPage = ConvertFrom-Html -URI $strNextDownloadPageURL
            $arrNodes = @($HtmlNodeDownloadPage.SelectNodes('//a[@href]') | Where-Object { $_.InnerText.ToLower() -eq 'mame_2010_renameset_prerequisites.csv' })
            if ($arrNodes.Count -eq 0) {
                Write-Error ('Failed to download the MAME 2010 RenameSet prerequisites file. Please download the file that looks like MAME_2010_RenameSet_Prerequisites.csv the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strPrerequisiteDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalPrerequisiteCSVFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalPrerequisiteCSVFilePath to point to the path of the downloaded CSV file.')
                break
            }
            $strNextURL = $arrNodes[0].Attributes['href'].Value

            $strURLBase = $strNextDownloadPageURL
            $strURLRelative = $strNextURL
            $strNextURL = Get-AbsoluteURLFromRelative $strURLBase $strURLRelative

            $strNextDownloadPageURL = $strNextURL
            $HtmlNodeDownloadPage = ConvertFrom-Html -URI $strNextDownloadPageURL
            $arrNodes = @($HtmlNodeDownloadPage.SelectNodes('//a[@href]') | Where-Object { $_.InnerText.ToLower() -eq 'Raw' })
            if ($arrNodes.Count -eq 0) {
                Write-Error ('Failed to download the MAME 2010 RenameSet prerequisites file. Please download the file that looks like MAME_2010_RenameSet_Prerequisites.csv the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strPrerequisiteDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalPrerequisiteCSVFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalPrerequisiteCSVFilePath to point to the path of the downloaded CSV file.')
                break
            }
            $strNextURL = $arrNodes[0].Attributes['href'].Value

            $strURLBase = $strNextDownloadPageURL
            $strURLRelative = $strNextURL
            $strNextURL = Get-AbsoluteURLFromRelative $strURLBase $strURLRelative

            $strEffectiveURL = $strNextURL
        }
        if ((Test-Path $strSubfolderPath) -ne $true) {
            New-Item $strSubfolderPath -ItemType Directory | Out-Null
        }
        Write-Verbose ('Downloading RenameSet prerequisites from ' + $strEffectiveURL + '...')
        $VerbosePreference = $actionPreferenceFormerVerbose
        Invoke-WebRequest -Uri $strEffectiveURL -OutFile (Join-Path $strSubfolderPath 'MAME_2010_RenameSet_Prerequisites.csv')
        $VerbosePreference = $actionPreferenceNewVerbose

        if (Test-Path (Join-Path $strSubfolderPath 'MAME_2010_RenameSet_Prerequisites.csv')) {
            # Successful download
            $arrRenameSetPrerequisites = @(Import-Csv -Path (Join-Path $strSubfolderPath 'MAME_2010_RenameSet_Prerequisites.csv'))
        } else {
            Write-Error ('Failed to download the MAME 2010 RenameSet prerequisites file. Please download the file that looks like MAME_2010_RenameSet_Prerequisites.csv the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strPrerequisiteDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalPrerequisiteCSVFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalPrerequisiteCSVFilePath to point to the path of the downloaded CSV file.')
            break
        }
    } else {
        if ((Test-Path $strLocalPrerequisiteCSVFilePath) -ne $true) {
            Write-Error ('The MAME 2010 RenameSet prerequisites file is missing. Please download the file that looks like MAME_2010_RenameSet_Prerequisites.csv from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strPrerequisiteDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalPrerequisiteCSVFilePath)
            break
        }
        $arrRenameSetPrerequisites = @(Import-Csv -Path (Join-Path $strSubfolderPath 'MAME_2010_RenameSet_Prerequisites.csv'))
    }
}

# Get the exclusion CSV
if ($null -eq $strLocalExclusionCSVFilePath -and $null -eq $strExclusionDownloadPageURL -and $null -eq $strExclusionCSVURL) {
    $arrRenameSetExclusions = @()
} else {
    $arrCommands = @(Get-Command Invoke-WebRequest)
    $boolInvokeWebRequestAvailable = ($arrCommands.Count -ge 1)
    if ($null -eq $strLocalExclusionCSVFilePath -and $boolInvokeWebRequestAvailable) {
        $VerbosePreference = $actionPreferenceFormerVerbose
        $arrModules = @(Get-Module PowerHTML -ListAvailable)
        $VerbosePreference = $actionPreferenceNewVerbose
        if ($arrModules.Count -eq 0) {
            Write-Warning 'It is recommended that you install the PowerHTML module using "Install-Module PowerHTML" before continuing. Doing so will allow this script to obtain the URL for the most-current RenameSet exclusion file automatically. Without PowerHTML, this script is using a potentially-outdated URL. Break out of ths script now to install PowerHTML, then re-run the script'
            $strEffectiveURL = $strExclusionCSVURL
        } else {
            Write-Verbose ('Parsing site ' + $strExclusionDownloadPageURL + ' to dynamically obtain RenameSet exclusion list download URL...')
            $arrLoadedModules = @(Get-Module PowerHTML)
            if ($arrLoadedModules.Count -eq 0) {
                $VerbosePreference = $actionPreferenceFormerVerbose
                Import-Module PowerHTML
                $VerbosePreference = $actionPreferenceNewVerbose
            }

            $strNextDownloadPageURL = $strExclusionDownloadPageURL
            $HtmlNodeDownloadPage = ConvertFrom-Html -URI $strNextDownloadPageURL
            $arrNodes = @($HtmlNodeDownloadPage.SelectNodes('//a[@href]') | Where-Object { $_.InnerText.ToLower() -eq 'mame_2010_renameset_exclusions.csv' })
            if ($arrNodes.Count -eq 0) {
                Write-Error ('Failed to download the MAME 2010 RenameSet exclusions file. Please download the file that looks like MAME_2010_RenameSet_Exclusions.csv the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strExclusionDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalExclusionCSVFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalExclusionCSVFilePath to point to the path of the downloaded CSV file.')
                break
            }
            $strNextURL = $arrNodes[0].Attributes['href'].Value

            $strURLBase = $strNextDownloadPageURL
            $strURLRelative = $strNextURL
            $strNextURL = Get-AbsoluteURLFromRelative $strURLBase $strURLRelative

            $strNextDownloadPageURL = $strNextURL
            $HtmlNodeDownloadPage = ConvertFrom-Html -URI $strNextDownloadPageURL
            $arrNodes = @($HtmlNodeDownloadPage.SelectNodes('//a[@href]') | Where-Object { $_.InnerText.ToLower() -eq 'Raw' })
            if ($arrNodes.Count -eq 0) {
                Write-Error ('Failed to download the MAME 2010 RenameSet exclusions file. Please download the file that looks like MAME_2010_RenameSet_Exclusions.csv the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strExclusionDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalExclusionCSVFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalExclusionCSVFilePath to point to the path of the downloaded CSV file.')
                break
            }
            $strNextURL = $arrNodes[0].Attributes['href'].Value

            $strURLBase = $strNextDownloadPageURL
            $strURLRelative = $strNextURL
            $strNextURL = Get-AbsoluteURLFromRelative $strURLBase $strURLRelative

            $strEffectiveURL = $strNextURL
        }
        if ((Test-Path $strSubfolderPath) -ne $true) {
            New-Item $strSubfolderPath -ItemType Directory | Out-Null
        }
        Write-Verbose ('Downloading RenameSet exclusion list from ' + $strEffectiveURL + '...')
        $VerbosePreference = $actionPreferenceFormerVerbose
        Invoke-WebRequest -Uri $strEffectiveURL -OutFile (Join-Path $strSubfolderPath 'MAME_2010_RenameSet_Exclusions.csv')
        $VerbosePreference = $actionPreferenceNewVerbose

        if (Test-Path (Join-Path $strSubfolderPath 'MAME_2010_RenameSet_Exclusions.csv')) {
            # Successful download
            $arrRenameSetExclusions = @(Import-Csv -Path (Join-Path $strSubfolderPath 'MAME_2010_RenameSet_Exclusions.csv'))
        } else {
            Write-Error ('Failed to download the MAME 2010 RenameSet exclusions file. Please download the file that looks like MAME_2010_RenameSet_Exclusions.csv the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strExclusionDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalExclusionCSVFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalExclusionCSVFilePath to point to the path of the downloaded CSV file.')
            break
        }
    } else {
        if ((Test-Path $strLocalExclusionCSVFilePath) -ne $true) {
            Write-Error ('The MAME 2010 RenameSet exclusions file is missing. Please download the file that looks like MAME_2010_RenameSet_Exclusions.csv from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strExclusionDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalExclusionCSVFilePath)
            break
        }
        $arrRenameSetExclusions = @(Import-Csv -Path (Join-Path $strSubfolderPath 'MAME_2010_RenameSet_Exclusions.csv'))
    }
}

# We have all the files, let's do stuff

$versionConvertedMAMEVersion = Convert-MAMEVersionNumberToRepresentativePowerShellVersion $strSourceMAMEVersion

$arrFilteredRenameSetExclusions = @($arrRenameSetExclusions | ForEach-Object {
        $PSCustomObjectRenameSetExclusionItem = $_
        $PSCustomObjectRenameSetExclusionItem.MAMEVersionPowerShellFriendly = [System.Version]($PSCustomObjectRenameSetExclusionItem.MAMEVersionPowerShellFriendly)
        $PSCustomObjectRenameSetExclusionItem
    } | Where-Object { $_.MAMEVersionPowerShellFriendly -gt $versionConvertedMAMEVersion })

Write-Verbose "Loading emulator ROM set's ROM package metadata..."
$arrROMPackageMetadata = Import-Csv $strPathToROMPackageMetadataCSV
$hashtableROMPackageMetadata = New-BackwardCompatibleCaseInsensitiveHashtable
$arrROMPackageMetadata |
    ForEach-Object {
        $PSCustomObjectROMMetadata = $_
        $hashtableROMPackageMetadata.Add($PSCustomObjectROMMetadata.ROMName, $PSCustomObjectROMMetadata)
    }

Write-Verbose "Loading RenameSet tabulated data..."
$arrProgettoSnapsRenameSet = Import-Csv $strPathToProgettoSnapsRenameSetInCSVFormat

Write-Verbose "Filtering RenameSet down to just the data that needs to be processed for this emulator version/ROM set..."
$arrFilteredProgettoSnapsRenameSet = @($arrProgettoSnapsRenameSet | ForEach-Object {
        $PSCustomObjectRenameSetItem = $_
        $PSCustomObjectRenameSetItem.MAMEVersionPowerShellFriendly = [System.Version]($PSCustomObjectRenameSetItem.MAMEVersionPowerShellFriendly)
        $PSCustomObjectRenameSetItem
    } | Where-Object { $_.MAMEVersionPowerShellFriendly -gt $versionConvertedMAMEVersion })

$arrConvertedMAMEVersionsToProcess = @($arrFilteredProgettoSnapsRenameSet |
        Select-Object -Property @('MAMEVersionPowerShellFriendly') -Unique) |
        ForEach-Object { $_.MAMEVersionPowerShellFriendly }

$intTotalRenameSetOperations = ($arrRenameSetPrerequisites.Count) + ($arrFilteredProgettoSnapsRenameSet.Count)
$intCurrentRenameSetOperation = 1
$timeDateStartOfProcessing = Get-Date

Write-Verbose "Applying RenameSet to this emulator version/ROM set..."
$strDisplayVersion = "Prerequisites"
$arrRenameSetPrerequisites |
    ForEach-Object {
        $PSCustomObjectRenameSetPrerequisiteItem = $_

        if ($intCurrentRenameSetOperation -ge 101) {
            $timeDateCurrent = Get-Date
            $timeSpanElapsed = $timeDateCurrent - $timeDateStartOfProcessing
            $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / ($intCurrentRenameSetOperation - 1) * $intTotalRenameSetOperations
            $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
            $doublePercentComplete = ($intCurrentRenameSetOperation - 1) / $intTotalRenameSetOperations * 100
            Write-Progress -Activity 'Applying RenameSet to This Emulator Version/ROM Set' -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
        }

        $strOldROMPackageName = $PSCustomObjectRenameSetPrerequisiteItem.OldROMPackageName
        $strNewROMPackageName = $PSCustomObjectRenameSetPrerequisiteItem.NewROMPackageName
        if ($PSCustomObjectRenameSetPrerequisiteItem.Operation -eq 'D') {
            # Delete operation
            if ($hashtableROMPackageMetadata.ContainsKey($strOldROMPackageName)) {
                $PSCustomObjectROMMetadata = $hashtableROMPackageMetadata.Item($strOldROMPackageName)
                $hashtableROMPackageMetadata.Remove($strOldROMPackageName)
                $PSCustomObjectROMMetadata.ROMName = ('Deleted*' + $strOldROMPackageName + '*' + $strDisplayVersion)
                $hashtableROMPackageMetadata.Add($PSCustomObjectROMMetadata.ROMName, $PSCustomObjectROMMetadata)
            } else {
                Write-Debug -Message ('RenameSet input file indicated that in version ' + $strDisplayVersion + ', MAME deleted the ROM "' + $strOldROMPackageName + '" - however, this ROM was not present in the ROM set as it was in the process of advancing to this version. Either the RenameSet information is wrong, or the ROM package was missing in the emulator metadata.')
            }
        } elseif ($PSCustomObjectRenameSetPrerequisiteItem.Operation -eq 'R') {
            # Rename operation
            if ($hashtableROMPackageMetadata.ContainsKey($strOldROMPackageName)) {
                $PSCustomObjectROMMetadata = $hashtableROMPackageMetadata.Item($strOldROMPackageName)
                $hashtableROMPackageMetadata.Remove($strOldROMPackageName)
                $PSCustomObjectROMMetadata.ROMName = ('Renamed*' + $strDisplayVersion + '*' + $strOldROMPackageName + '*' + $strNewROMPackageName)
                $hashtableROMPackageMetadata.Add($PSCustomObjectROMMetadata.ROMName, $PSCustomObjectROMMetadata)
            } else {
                Write-Debug -Message ('RenameSet input file indicated that in version ' + $strDisplayVersion + ', MAME renamed the ROM "' + $strOldROMPackageName + '" to something new - however, this ROM was not present in the ROM set as it was in the process of advancing to this version. Either the RenameSet information is wrong, or the ROM package was missing in the emulator metadata.')
            }
        }

        $intCurrentRenameSetOperation++
    }
# Process pending renames
$strPrefix = 'Renamed*' + $strDisplayVersion + '*'
@($hashtableROMPackageMetadata.Keys | Where-Object { $_.Contains($strPrefix) }) |
    ForEach-Object {
        $strTemporaryROMNameToBeRenamed = $_
        $arrTemporaryROMNameToBeRenamed = Split-StringOnLiteralString $strTemporaryROMNameToBeRenamed $strPrefix
        $strOldAndNewROMPackageName = $arrTemporaryROMNameToBeRenamed[$arrTemporaryROMNameToBeRenamed.Length - 1]
        $arrOldAndNewROMPackageNames = Split-StringOnLiteralString $strOldAndNewROMPackageName '*'
        $strOldROMPackageName = $arrOldAndNewROMPackageNames[0]
        $strNewROMPackageName = $arrOldAndNewROMPackageNames[$arrOldAndNewROMPackageNames.Length - 1]
        $PSCustomObjectROMMetadata = $hashtableROMPackageMetadata.Item($strTemporaryROMNameToBeRenamed)
        if ($hashtableROMPackageMetadata.ContainsKey($strNewROMPackageName)) {
            # Collision
            $strCollisionName = 'RenameSetCollision*' + $strOldROMPackageName + '*to*' + $strNewROMPackageName + '*AtVersion*' + $strDisplayVersion
            Write-Warning -Message ('RenameSet input file indicated that in version ' + $strDisplayVersion + ', MAME renamed the ROM "' + $strOldROMPackageName + '" to "' + $strNewROMPackageName + ' - however, there is already a ROM present in this ROM set as it was in the process of advancing to this version with the name "' + $strNewROMPackageName + '". Either the RenameSet information is wrong, or the emulator was updated after the official release of the emulator upon which the RenameSet is based. The ROM package metadata will be retained with a ROMName "' + $strCollisionName + '" - it is recommended that you inspect the input and output to determine if this is expected.')
            $hashtableROMPackageMetadata.Remove($strTemporaryROMNameToBeRenamed)
            $PSCustomObjectROMMetadata.ROMName = $strCollisionName
            $hashtableROMPackageMetadata.Add($PSCustomObjectROMMetadata.ROMName, $PSCustomObjectROMMetadata)
        } else {
            # Happy path
            $hashtableROMPackageMetadata.Remove($strTemporaryROMNameToBeRenamed)
            $PSCustomObjectROMMetadata.ROMName = $strNewROMPackageName
            $hashtableROMPackageMetadata.Add($PSCustomObjectROMMetadata.ROMName, $PSCustomObjectROMMetadata)
        }
    }

$arrConvertedMAMEVersionsToProcess | ForEach-Object {
    $versionCurrentlyProcessing = $_
    $PSCustomObjectRenameSetItem = @($arrFilteredProgettoSnapsRenameSet | Where-Object { $_.MAMEVersionPowerShellFriendly -eq $versionCurrentlyProcessing } | Select-Object -First 1)[0]
    $strDisplayVersion = $PSCustomObjectRenameSetItem.MAMEVersion
    Write-Debug ('Processing emulator version ' + $strDisplayVersion + '...')
    $arrFilteredProgettoSnapsRenameSet | Where-Object { $_.MAMEVersionPowerShellFriendly -eq $versionCurrentlyProcessing } | Sort-Object -Property @('MAMEVersionPowerShellFriendly', 'MAMEDate', 'Operation', 'OldROMPackageName') | ForEach-Object {
        $PSCustomObjectRenameSetItem = $_

        if ($intCurrentRenameSetOperation -ge 101) {
            $timeDateCurrent = Get-Date
            $timeSpanElapsed = $timeDateCurrent - $timeDateStartOfProcessing
            $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / ($intCurrentRenameSetOperation - 1) * $intTotalRenameSetOperations
            $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
            $doublePercentComplete = ($intCurrentRenameSetOperation - 1) / $intTotalRenameSetOperations * 100
            Write-Progress -Activity 'Applying RenameSet to This Emulator Version/ROM Set' -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
        }

        if ((Test-RenameSetRecordShouldBeIgnored ([ref]$arrFilteredRenameSetExclusions) ([ref]$PSCustomObjectRenameSetItem)) -ne $true) {
            $strOldROMPackageName = $PSCustomObjectRenameSetItem.OldROMPackageName
            $strNewROMPackageName = $PSCustomObjectRenameSetItem.NewROMPackageName
            if ($PSCustomObjectRenameSetItem.Operation -eq 'D') {
                # Delete operation
                if ($hashtableROMPackageMetadata.ContainsKey($strOldROMPackageName)) {
                    $PSCustomObjectROMMetadata = $hashtableROMPackageMetadata.Item($strOldROMPackageName)
                    $hashtableROMPackageMetadata.Remove($strOldROMPackageName)
                    $PSCustomObjectROMMetadata.ROMName = ('Deleted*' + $strOldROMPackageName + '*' + $strDisplayVersion)
                    $hashtableROMPackageMetadata.Add($PSCustomObjectROMMetadata.ROMName, $PSCustomObjectROMMetadata)
                } else {
                    Write-Debug -Message ('RenameSet input file indicated that in version ' + $strDisplayVersion + ', MAME deleted the ROM "' + $strOldROMPackageName + '" - however, this ROM was not present in the ROM set as it was in the process of advancing to this version. Either the RenameSet information is wrong, or the ROM package was missing in the emulator metadata.')
                }
            } elseif ($PSCustomObjectRenameSetItem.Operation -eq 'R') {
                # Rename operation
                if ($hashtableROMPackageMetadata.ContainsKey($strOldROMPackageName)) {
                    $PSCustomObjectROMMetadata = $hashtableROMPackageMetadata.Item($strOldROMPackageName)
                    $hashtableROMPackageMetadata.Remove($strOldROMPackageName)
                    $PSCustomObjectROMMetadata.ROMName = ('Renamed*' + $strDisplayVersion + '*' + $strOldROMPackageName + '*' + $strNewROMPackageName)
                    $hashtableROMPackageMetadata.Add($PSCustomObjectROMMetadata.ROMName, $PSCustomObjectROMMetadata)
                } else {
                    Write-Debug -Message ('RenameSet input file indicated that in version ' + $strDisplayVersion + ', MAME renamed the ROM "' + $strOldROMPackageName + '" to something new - however, this ROM was not present in the ROM set as it was in the process of advancing to this version. Either the RenameSet information is wrong, or the ROM package was missing in the emulator metadata.')
                }
            }
        } else {
            Write-Debug -Message ('RenameSet input file indicated that in version ' + $strDisplayVersion + ', MAME renamed the ROM "' + $strOldROMPackageName + '" to "' + $strNewROMPackageName + '" - however, this operation is on the ignore list and will not be processed.')
        }

        $intCurrentRenameSetOperation++
    }
    # Process pending renames
    $strPrefix = 'Renamed*' + $strDisplayVersion + '*'
    @($hashtableROMPackageMetadata.Keys | Where-Object { $_.Contains($strPrefix) }) | ForEach-Object {
        $strTemporaryROMNameToBeRenamed = $_
        $arrTemporaryROMNameToBeRenamed = Split-StringOnLiteralString $strTemporaryROMNameToBeRenamed $strPrefix
        $strOldAndNewROMPackageName = $arrTemporaryROMNameToBeRenamed[$arrTemporaryROMNameToBeRenamed.Length - 1]
        $arrOldAndNewROMPackageNames = Split-StringOnLiteralString $strOldAndNewROMPackageName '*'
        $strOldROMPackageName = $arrOldAndNewROMPackageNames[0]
        $strNewROMPackageName = $arrOldAndNewROMPackageNames[$arrOldAndNewROMPackageNames.Length - 1]
        $PSCustomObjectROMMetadata = $hashtableROMPackageMetadata.Item($strTemporaryROMNameToBeRenamed)
        if ($hashtableROMPackageMetadata.ContainsKey($strNewROMPackageName)) {
            # Collision
            $strCollisionName = 'RenameSetCollision*' + $strOldROMPackageName + '*to*' + $strNewROMPackageName + '*AtVersion*' + $strDisplayVersion
            Write-Warning -Message ('RenameSet input file indicated that in version ' + $strDisplayVersion + ', MAME renamed the ROM "' + $strOldROMPackageName + '" to "' + $strNewROMPackageName + ' - however, there is already a ROM present in this ROM set as it was in the process of advancing to this version with the name "' + $strNewROMPackageName + '". Either the RenameSet information is wrong, or the emulator was updated after the official release of the emulator upon which the RenameSet is based. The ROM package metadata will be retained with a ROMName "' + $strCollisionName + '" - it is recommended that you inspect the input and output to determine if this is expected.')
            $hashtableROMPackageMetadata.Remove($strTemporaryROMNameToBeRenamed)
            $PSCustomObjectROMMetadata.ROMName = $strCollisionName
            $hashtableROMPackageMetadata.Add($PSCustomObjectROMMetadata.ROMName, $PSCustomObjectROMMetadata)
        } else {
            # Happy path
            $hashtableROMPackageMetadata.Remove($strTemporaryROMNameToBeRenamed)
            $PSCustomObjectROMMetadata.ROMName = $strNewROMPackageName
            $hashtableROMPackageMetadata.Add($PSCustomObjectROMMetadata.ROMName, $PSCustomObjectROMMetadata)
        }
    }
}

Write-Verbose ('Exporting results to CSV: ' + $strCSVOutputFile)
$hashtableROMPackageMetadata.Values | Sort-Object -Property @('ROMName') |
    Export-Csv -Path $strCSVOutputFile -NoTypeInformation

$VerbosePreference = $actionPreferenceFormerVerbose
$DebugPreference = $actionPreferenceFormerDebug
