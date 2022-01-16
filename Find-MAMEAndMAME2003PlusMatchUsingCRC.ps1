# Find-MAMEAndMAME2003PlusMatchUsingCRC.ps1

$strThisScriptVersionNumber = [version]'2.2.20220115.0'

#region License
###############################################################################################
# Copyright 2022 Frank Lesniak

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
#endregion

$actionPreferenceFormerVerbose = $VerbosePreference
$actionPreferenceFormerDebug = $DebugPreference

$strLocalMAMEDATWithCRCInfoCSV = $null
$strScriptToGenerateMAMEDATCSV = 'coresponding "Convert-..."'
$strMAMEDATDisplayName = 'the MAME DAT'
$strMAMEDATColumnPrefix = $null
$strMAMEDATROMNameColumnHeader = $null
$strMAMEDATROMDisplaNameColumnHeader = $null
$strMAMEDATROMFileCRCsColumnHeader = $null
$strMAMEDATCSVFilePrefix = $null
$strLocalDATToCompareToMAMEIncludingCRC = $null
$strLocalDATToCompareToMAME = $null
$strLocalDATToCompareToMAMEPrimaryKeyColumnHeader = $null
$strScriptToGenerateLocalDATToCompareToMAMEIncludingCRC = 'coresponding "Convert-..."'
$strLocalDATToCompareToMAMEDisplayName = 'this DAT to compare to MAME'
$strLocalDATToCompareToMAMEColumnPrefix = $null
$strLocalDATToCompareToMAMEROMNameColumnHeader = $null
$strLocalDATToCompareToMAMEROMDisplaNameColumnHeader = $null
$strLocalDATToCompareToMAMEROMFileCRCsColumnHeader = $null
$strLocalDATToCompareToMAMECSVFilePrefix = $null
$strROMFileCRCSeparator = '`t'
$strCSVOutputFileMatchedROMInfo = $null
$boolAlwaysAppendUnmatchedROMWithLocalDATToCompareToMAMEColumnPrefix = $false
$strCSVOutputFileRenamedROMDAT = $null
$actionPreferenceNewVerbose = $VerbosePreference
$actionPreferenceNewDebug = $DebugPreference

#region Inputs
###############################################################################################
# This script requires the current version of MAME's DAT converted to with alphabetized CRC
# hashes. Set the following path to point to the corresponding CSV:
$strLocalMAMEDATWithCRCInfoCSV = Join-Path '.' 'MAME_DAT_ROM_File_CRCs.csv'

# If the above CSV is missing, the script will tell the user to go generate it. Adjust the
# following to match the name of the correct script that the user should be instructed to run:
$strScriptToGenerateMAMEDATCSV = '"Convert-MAMEDATToCsv"'

# This display name is used in progress output:
$strMAMEDATDisplayName = "the MAME DAT's ROM CRCs"

$strMAMEDATColumnPrefix = 'MAME'
$strMAMEDATROMNameColumnHeader = 'MAME_ROMName'
$strMAMEDATROMDisplaNameColumnHeader = 'MAME_ROMDisplayName'
$strMAMEDATROMFileCRCsColumnHeader = 'MAME_ROMFileString'
$strMAMEDATCSVFilePrefix = 'MAME_DAT'

# This script also requires another CSV including alphabetized CRC information, which will be
# compared to the MAME DAT. Set the following path to point to the corresponding CSV:
$strLocalDATToCompareToMAMEIncludingCRC = Join-Path '.' 'MAME_2003_Plus_DAT_ROM_File_CRCs.csv'

# This script requires a DAT converted to CSV, which will be used to perform the "rename"
# operation that changes ROM names to match MAME
$strLocalDATToCompareToMAME = Join-Path '.' 'MAME_2003_Plus_DAT.csv'

$strLocalDATToCompareToMAMEPrimaryKeyColumnHeader = 'ROMName'

# If the above CSV is missing, the script will tell the user to go generate it. Adjust the
# following to match the name of the correct script that the user should be instructed to run:
$strScriptToGenerateLocalDATToCompareToMAMEIncludingCRC = '"Convert-MAME2003PlusDATToCSV"'

# This display name is used in progress output:
$strLocalDATToCompareToMAMEDisplayName = "the MAME 2003 Plus DAT's ROM CRCs"

$strLocalDATToCompareToMAMEColumnPrefix = 'MAME2003Plus'
$strLocalDATToCompareToMAMEROMNameColumnHeader = 'MAME2003Plus_ROMName'
$strLocalDATToCompareToMAMEROMDisplaNameColumnHeader = 'MAME2003Plus_ROMDisplayName'
$strLocalDATToCompareToMAMEROMFileCRCsColumnHeader = 'MAME2003Plus_ROMFileString'
$strLocalDATToCompareToMAMECSVFilePrefix = 'MAME_2003_Plus_DAT'

# If the input files use a different separator between ROM file CRC hashes, change this
$strROMFileCRCSeparator = "`t"

$strCSVOutputFileMatchedROMInfo = Join-Path '.' ($strLocalDATToCompareToMAMECSVFilePrefix + '_To_' + $strMAMEDATCSVFilePrefix + '_Mapping.csv')

$boolAlwaysAppendUnmatchedROMWithLocalDATToCompareToMAMEColumnPrefix = $false

$strCSVOutputFileRenamedROMDAT = Join-Path '.' ($strLocalDATToCompareToMAMECSVFilePrefix + '_Renamed_and_CRC-Matched_To_' + $strMAMEDATCSVFilePrefix + '.csv')

# Comment-out the following line if you prefer that the script operate silently.
$actionPreferenceNewVerbose = [System.Management.Automation.ActionPreference]::Continue

# Remove the comment from the following line if you prefer that the script output extra
# debugging information.
# $actionPreferenceNewDebug = [System.Management.Automation.ActionPreference]::Continue

###############################################################################################
#endregion Inputs

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

$VerbosePreference = $actionPreferenceNewVerbose
$DebugPreference = $actionPreferenceNewDebug

$boolErrorOccurred = $false

$dateTimeStart = Get-Date

if ((Test-Path $strLocalMAMEDATWithCRCInfoCSV) -ne $true) {
    Write-Error ('The input file "' + $strLocalMAMEDATWithCRCInfoCSV + '" is missing. Please generate it using the ' + $strScriptToGenerateMAMEDATCSV + ' script and then re-run this script')
    $boolErrorOccurred = $true
}

if ((Test-Path $strLocalDATToCompareToMAMEIncludingCRC) -ne $true) {
    Write-Error ('The input file "' + $strLocalDATToCompareToMAMEIncludingCRC + '" is missing. Please generate it using the ' + $strScriptToGenerateLocalDATToCompareToMAMEIncludingCRC + ' script and then re-run this script')
    $boolErrorOccurred = $true
}

if ((Test-Path $strLocalDATToCompareToMAME) -ne $true) {
    ('The input file "' + $strLocalDATToCompareToMAME + '" is missing. Please generate it using the ' + $strScriptToGenerateLocalDATToCompareToMAMEIncludingCRC + ' script and then re-run this script')
    $boolErrorOccurred = $true
}

$VerbosePreference = $actionPreferenceFormerVerbose
$arrModules = @(Get-Module Communary.PASM -ListAvailable)
$VerbosePreference = $actionPreferenceNewVerbose

if ($arrModules.Count -eq 0) {
    Write-Error 'This script requires the module "Communary.PASM". On PowerShell version 5.0 and newer, it can be instaled using the command "Install-Module Communary.PASM". Please install it and re-run the script'
    $boolErrorOccurred = $true
}

if ($boolErrorOccurred -eq $true) {
    break
}

Write-Verbose ('Importing ' + $strMAMEDATDisplayName + '...')
$arrMAMEDATWithCRCInfo = @(Import-Csv -Path $strLocalMAMEDATWithCRCInfoCSV)

Write-Verbose ('Importing ' + $strLocalDATToCompareToMAMEDisplayName + '...')
$arrLocalDATToCompareToMAMEWithCRCInfo = @(Import-Csv -Path $strLocalDATToCompareToMAMEIncludingCRC)

Write-Verbose ('Importing ' + $strLocalDATToCompareToMAME + '...')
$arrLocalDATToCompareToMAME = @(Import-Csv -Path $strLocalDATToCompareToMAME)

$arrLoadedModules = @(Get-Module Communary.PASM)
if ($arrLoadedModules.Count -eq 0) {
    $VerbosePreference = $actionPreferenceFormerVerbose
    Import-Module Communary.PASM
    $VerbosePreference = $actionPreferenceNewVerbose
}

# Build hashtables of MAME DAT for rapid lookup
$timeDateStartOfProcessing = Get-Date
$intTotalROMPackages = $arrMAMEDATWithCRCInfo.Count
$hashtableMAMEROMFileCRCsToROMNames = New-BackwardCompatibleCaseInsensitiveHashtable
$hashtableMAMEROMNamesToROMFileCount = New-BackwardCompatibleCaseInsensitiveHashtable
for ($intCounter = 0; $intCounter -lt $intTotalROMPackages; $intCounter++) {
    if ($intCounter -ge 101) {
        $timeDateCurrent = Get-Date
        $timeSpanElapsed = $timeDateCurrent - $timeDateStartOfProcessing
        $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / $intCounter * $intTotalROMPackages
        $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
        $doublePercentComplete = $intCounter / $intTotalROMPackages * 100
        Write-Progress -Activity ('Building ' + $strMAMEDATDisplayName + ' hashtables...') -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
    }

    $arrMAMEProperties = @(($arrMAMEDATWithCRCInfo[$intCounter]).PSObject.Properties)
    $strMAMEROMName = ($arrMAMEProperties | Where-Object { $_.Name -eq $strMAMEDATROMNameColumnHeader }).Value
    $strMAMEROMFileCRCs = ($arrMAMEProperties | Where-Object { $_.Name -eq $strMAMEDATROMFileCRCsColumnHeader }).Value

    if ([string]::IsNullOrEmpty($strMAMEROMFileCRCs) -eq $false) {
        $arrROMFileHashes = Split-StringOnLiteralString $strMAMEROMFileCRCs $strROMFileCRCSeparator
        $arrROMFileHashes | ForEach-Object {
            $strThisHash = $_
            if ($hashtableMAMEROMFileCRCsToROMNames.ContainsKey($strThisHash) -eq $true) {
                # Another ROM had this same ROM file
                $hashtableMAMEROMFileCRCsToROMNames.Item($strThisHash) += $strMAMEROMName
            } else {
                $hashtableMAMEROMFileCRCsToROMNames.Add($strThisHash, @($strMAMEROMName))
            }
        }
    } else {
        $arrROMFileHashes = @()
    }
    if ($hashtableMAMEROMNamesToROMFileCount.ContainsKey($strMAMEROMName) -eq $false) {
        $hashtableMAMEROMNamesToROMFileCount.Add($strMAMEROMName, $arrROMFileHashes.Count)
    } else {
        Write-Warning ('The ' + $strMAMEDATDisplayName + ' contained multiple rows for the machine/game named ' + $strMAMEROMName + '. This should not be possible and indicates a problem with the source data.')
    }
}

# Compare local DAT to MAME
$timeDateStartOfProcessing = Get-Date
$intTotalROMPackages = $arrLocalDATToCompareToMAMEWithCRCInfo.Count
$arrOutput = @()
for ($intCounter = 0; $intCounter -lt $intTotalROMPackages; $intCounter++) {
    if ($intCounter -ge 101) {
        $timeDateCurrent = Get-Date
        $timeSpanElapsed = $timeDateCurrent - $timeDateStartOfProcessing
        $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / $intCounter * $intTotalROMPackages
        $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
        $doublePercentComplete = $intCounter / $intTotalROMPackages * 100
        Write-Progress -Activity ('Comparing ' + $strLocalDATToCompareToMAMEDisplayName + ' to ' + $strMAMEDATDisplayName + '...') -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
    }

    $hashtableMatchedROMs = New-BackwardCompatibleCaseInsensitiveHashtable

    $arrMachineToCompareToMAMEProperties = @(($arrLocalDATToCompareToMAMEWithCRCInfo[$intCounter]).PSObject.Properties)
    $strMachineToCompareToMAMEROMName = ($arrMachineToCompareToMAMEProperties | Where-Object { $_.Name -eq $strLocalDATToCompareToMAMEROMNameColumnHeader }).Value
    $strMachineToCompareToMAMEROMDisplayName = ($arrMachineToCompareToMAMEProperties | Where-Object { $_.Name -eq $strLocalDATToCompareToMAMEROMDisplaNameColumnHeader }).Value
    $strMachineToCompareToMAMEROMFileCRCs = ($arrMachineToCompareToMAMEProperties | Where-Object { $_.Name -eq $strLocalDATToCompareToMAMEROMFileCRCsColumnHeader }).Value

    if ([string]::IsNullOrEmpty($strMachineToCompareToMAMEROMFileCRCs) -eq $false) {
        $arrROMFileHashes = Split-StringOnLiteralString $strMachineToCompareToMAMEROMFileCRCs $strROMFileCRCSeparator
        $arrROMFileHashes | ForEach-Object {
            $strThisHash = $_
            if ($hashtableMAMEROMFileCRCsToROMNames.ContainsKey($strThisHash) -eq $true) {
                $arrMatchedROMNames = @($hashtableMAMEROMFileCRCsToROMNames.Item($strThisHash))
                $arrMatchedROMNames | ForEach-Object {
                    $strThisMatchedROM = $_
                    if ($hashtableMatchedROMs.ContainsKey($strThisMatchedROM) -eq $true) {
                        $hashtableMatchedROMs.Item($strThisMatchedROM)++
                    } else {
                        $hashtableMatchedROMs.Add($strThisMatchedROM, 1)
                    }
                }
            } else {
                # Write-Host ('No match for: ' + $strThisHash)
            }
        }
    } else {
        $arrROMFileHashes = @()
    }

    $arrTopMatches = @($hashtableMatchedROMs.GetEnumerator() | Sort-Object -Property 'Value' -Descending |
            Select-Object -First 25)

    $arrRevisedTopMatches = $arrTopMatches | ForEach-Object {
        $strMatchedROMName = $_.Key
        $intNumMatchedROMsFromPerspectiveOfMachineToCompareToMAME = $_.Value
        $intNumROMFilesToMatchFromPerspectiveOfMachineToCompareToMAME = $arrROMFileHashes.Count
        $intNumROMFilesToMatchFromPerspectiveOfMAME = $hashtableMAMEROMNamesToROMFileCount.Item($strMatchedROMName)
        if ($intNumROMFilesToMatchFromPerspectiveOfMachineToCompareToMAME -ne 0) {
            $doublePercentMatchFromPerspectiveOfMachineToCompareToMAME = $intNumMatchedROMsFromPerspectiveOfMachineToCompareToMAME / $intNumROMFilesToMatchFromPerspectiveOfMachineToCompareToMAME
        } else {
            $doublePercentMatchFromPerspectiveOfMachineToCompareToMAME = 0
        }
        if ($intNumROMFilesToMatchFromPerspectiveOfMAME -ne 0) {
            $doublePercentMatchFromPerspectiveOfMAME = $intNumMatchedROMsFromPerspectiveOfMachineToCompareToMAME / $intNumROMFilesToMatchFromPerspectiveOfMAME
        } else {
            $doublePercentMatchFromPerspectiveOfMAME = 0
        }
        $doubleAveragePercentMatch = ($doublePercentMatchFromPerspectiveOfMachineToCompareToMAME + $doublePercentMatchFromPerspectiveOfMAME) / 2
        $doubleMatchedROMNamePercentSimilarity = (Get-PasmScore -String1 $strMachineToCompareToMAMEROMName -String2 $strMatchedROMName -CaseSensitive:$false -Algorithm LevenshteinDistance) / 100

        $PSObjectRevisedTopMatch = New-Object PSObject
        $PSObjectRevisedTopMatch | Add-Member -MemberType NoteProperty -Name 'MatchedROMName' -Value $strMatchedROMName
        $PSObjectRevisedTopMatch | Add-Member -MemberType NoteProperty -Name 'AvgPercentMatch' -Value $doubleAveragePercentMatch
        $PSObjectRevisedTopMatch | Add-Member -MemberType NoteProperty -Name 'PercentMatchROMName' -Value $doubleMatchedROMNamePercentSimilarity
        $PSObjectRevisedTopMatch | Add-Member -MemberType NoteProperty -Name 'PercentMatchFromPerspectiveOfMachineToCompareToMAME' -Value $doublePercentMatchFromPerspectiveOfMachineToCompareToMAME
        $PSObjectRevisedTopMatch | Add-Member -MemberType NoteProperty -Name 'PercentMatchFromPerspectiveOfMAME' -Value $doublePercentMatchFromPerspectiveOfMAME

        $PSObjectRevisedTopMatch
    } | Sort-Object -Property @('AvgPercentMatch', 'PercentMatchROMName') -Descending

    $PSObjectMatches = New-Object PSObject
    $PSObjectMatches | Add-Member -MemberType NoteProperty -Name $strLocalDATToCompareToMAMEROMNameColumnHeader -Value $strMachineToCompareToMAMEROMName
    $PSObjectMatches | Add-Member -MemberType NoteProperty -Name $strLocalDATToCompareToMAMEROMDisplaNameColumnHeader -Value $strMachineToCompareToMAMEROMDisplayName

    for ($intInnerCounter = 0; $intInnerCounter -lt 25; $intInnerCounter++) {
        if ($intInnerCounter -lt $arrRevisedTopMatches.Count) {
            $strThisMatchedROMName = ($arrRevisedTopMatches[$intInnerCounter]).MatchedROMName
            $strAveragePercentMatch = [string](($arrRevisedTopMatches[$intInnerCounter]).AvgPercentMatch)
            $strPercentMatchFromPerspectiveOfMachineToCompareToMAME = [string](($arrRevisedTopMatches[$intInnerCounter]).PercentMatchFromPerspectiveOfMachineToCompareToMAME)
            $strPercentMatchFromPerspectiveOfMAME = [string](($arrRevisedTopMatches[$intInnerCounter]).PercentMatchFromPerspectiveOfMAME)
        } else {
            $strThisMatchedROMName = ''
            $strAveragePercentMatch = ''
            $strPercentMatchFromPerspectiveOfMachineToCompareToMAME = ''
            $strPercentMatchFromPerspectiveOfMAME = ''
        }
        $strColumnName = ($strLocalDATToCompareToMAMEColumnPrefix + '_MatchedTo' + $strMAMEDATColumnPrefix + '_' + ([string]($intInnerCounter + 1)) + '_ROMName')
        $PSObjectMatches | Add-Member -MemberType NoteProperty -Name $strColumnName -Value $strThisMatchedROMName
        $strColumnName = ($strLocalDATToCompareToMAMEColumnPrefix + '_MatchedTo' + $strMAMEDATColumnPrefix + '_' + ([string]($intInnerCounter + 1)) + '_' + 'AveragePercentMatch')
        $PSObjectMatches | Add-Member -MemberType NoteProperty -Name $strColumnName -Value $strAveragePercentMatch
        $strColumnName = ($strLocalDATToCompareToMAMEColumnPrefix + '_MatchedTo' + $strMAMEDATColumnPrefix + '_' + ([string]($intInnerCounter + 1)) + '_' + $strLocalDATToCompareToMAMEColumnPrefix + 'PerspectiveMatchPercentage')
        $PSObjectMatches | Add-Member -MemberType NoteProperty -Name $strColumnName -Value $strPercentMatchFromPerspectiveOfMachineToCompareToMAME
        $strColumnName = ($strLocalDATToCompareToMAMEColumnPrefix + '_MatchedTo' + $strMAMEDATColumnPrefix + '_' + ([string]($intInnerCounter + 1)) + '_' + $strMAMEDATColumnPrefix + 'PerspectiveMatchPercentage')
        $PSObjectMatches | Add-Member -MemberType NoteProperty -Name $strColumnName -Value $strPercentMatchFromPerspectiveOfMAME
    }

    $arrOutput += $PSObjectMatches
}

Write-Verbose ('Exporting matching results to CSV: ' + $strCSVOutputFileMatchedROMInfo)
$strColumnName = ($strLocalDATToCompareToMAMEColumnPrefix + '_MatchedTo' + $strMAMEDATColumnPrefix + '_1_AveragePercentMatch')
$strScriptblock = '$_.' + $strColumnName
$scriptblock = [scriptblock]::Create($strScriptblock)
$hashtableSortDescendingProperty = New-BackwardCompatibleCaseInsensitiveHashtable
$hashtableSortDescendingProperty.Add('Expression', $scriptblock)
$hashtableSortDescendingProperty.Add('Ascending', $false)
$strAscendingProperty = ($strLocalDATToCompareToMAMEColumnPrefix + '_MatchedTo' + $strMAMEDATColumnPrefix + '_1_ROMName')
$arrSortedOutput = $arrOutput | Sort-Object -Property @($strAscendingProperty, $hashtableSortDescendingProperty, $strLocalDATToCompareToMAMEROMNameColumnHeader)
$arrSortedOutput | Export-Csv -Path $strCSVOutputFileMatchedROMInfo -NoTypeInformation

$hashtableAllMappings = New-BackwardCompatibleCaseInsensitiveHashtable

$timeDateStartOfProcessing = Get-Date
$intTotalROMPackages = $arrSortedOutput.Count
$intCounter = 0
$arrSortedOutput | ForEach-Object {
    if ($intCounter -ge 101) {
        $timeDateCurrent = Get-Date
        $timeSpanElapsed = $timeDateCurrent - $timeDateStartOfProcessing
        $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / $intCounter * $intTotalROMPackages
        $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
        $doublePercentComplete = $intCounter / $intTotalROMPackages * 100
        Write-Progress -Activity ('Finalizing ' + $strLocalDATToCompareToMAMEDisplayName + ' to ' + $strMAMEDATDisplayName + '...') -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
    }

    $objThisMatchedROM = $_
    $arrMatchedROMProperties = @($objThisMatchedROM.PSObject.Properties)
    $strThisLocalDATROMName = ($arrMatchedROMProperties | Where-Object { $_.Name -eq $strLocalDATToCompareToMAMEROMNameColumnHeader }).Value
    $strColumnName = ($strLocalDATToCompareToMAMEColumnPrefix + '_MatchedTo' + $strMAMEDATColumnPrefix + '_1_ROMName')
    $strThisMAMEROMName = ($arrMatchedROMProperties | Where-Object { $_.Name -eq $strColumnName }).Value
    $strColumnName = ($strLocalDATToCompareToMAMEColumnPrefix + '_MatchedTo' + $strMAMEDATColumnPrefix + '_1_AveragePercentMatch')
    $doubleAveragePercentMatch = [double](($arrMatchedROMProperties | Where-Object { $_.Name -eq $strColumnName }).Value)
    
    $boolUnmatched = $true
    if ($doubleAveragePercentMatch -ge 0.5) {
        if ($hashtableAllMappings.ContainsValue($strThisMAMEROMName) -eq $true) {
            # This is normal/expected, so treat as unmatched
        } else {
            if ($hashtableAllMappings.ContainsKey($strThisLocalDATROMName) -eq $true) {
                # This is unexpected, throw a warning
                Write-Warning ('Somehow, our ROM mapping hashtable already contains a key (in a key-value pair) for ' + $strThisLocalDATROMName + '. So we will try treating it as an unmatched ROM')
            } else {
                $hashtableAllMappings.Add($strThisLocalDATROMName, $strThisMAMEROMName)
                $boolUnmatched = $false
            }
        }
    }

    if ($boolUnmatched -eq $true) {
        # We concluded that this ROM does not match a MAME ROM
        if ($hashtableMAMEROMNamesToROMFileCount.ContainsKey($strThisLocalDATROMName) -or $boolAlwaysAppendUnmatchedROMWithLocalDATToCompareToMAMEColumnPrefix -eq $true) {
            # ... but the MAME ROM set already contains a ROM with this name
            # so, let's append the name of the ROM set in front of it:
            $strUnmatchedROMName = $strLocalDATToCompareToMAMEColumnPrefix + '_' + $strThisLocalDATROMName
            if ($hashtableMAMEROMNamesToROMFileCount.ContainsKey($strUnmatchedROMName)) {
                Write-Warning ('Somehow, ' + $strMAMEDATDisplayName + ' already contains an entry for ' + $strUnmatchedROMName + ', which may cause an incorrect ROM match in a downstream process.')
            }
        } else {
            $strUnmatchedROMName = $strThisLocalDATROMName
        }

        if ($hashtableAllMappings.ContainsValue($strUnmatchedROMName) -eq $true) {
            Write-Warning ('Somehow, our ROM mapping hashtable already contains a value (in a key-value pair) for ' + $strUnmatchedROMName + '. So we are dropping an entry, which may cause an incomplete database')
        } else {
            if ($hashtableAllMappings.ContainsKey($strThisLocalDATROMName) -eq $true) {
                Write-Warning ('Somehow, our ROM mapping hashtable already contains a key (in a key-value pair) for ' + $strThisLocalDATROMName + '. So we are dropping an entry, which may cause an incomplete database')
            } else {
                $hashtableAllMappings.Add($strThisLocalDATROMName, $strUnmatchedROMName)
            }
        }
    }

    $intCounter++
}

$timeDateStartOfProcessing = Get-Date
$intTotalROMPackages = $arrLocalDATToCompareToMAME.Count
$intCounter = 0
Write-Verbose ('Applying ROM renames to ' + $strLocalDATToCompareToMAMEDisplayName + '...')
$arrRevisedDAT = $arrLocalDATToCompareToMAME | ForEach-Object {
    if ($intCounter -ge 101) {
        $timeDateCurrent = Get-Date
        $timeSpanElapsed = $timeDateCurrent - $timeDateStartOfProcessing
        $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / $intCounter * $intTotalROMPackages
        $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
        $doublePercentComplete = $intCounter / $intTotalROMPackages * 100
        Write-Progress -Activity ('Applying ROM renames to ' + $strLocalDATToCompareToMAMEDisplayName + '...') -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
    }
    $objThisMachineToCompareToMAME = $_
    $arrMachineProperties = @($objThisMachineToCompareToMAME.PSObject.Properties)
    $strThisLocalDATROMName = ($arrMachineProperties | Where-Object { $_.Name -eq $strLocalDATToCompareToMAMEROMNameColumnHeader }).Value
    if ($hashtableAllMappings.ContainsKey($strThisLocalDATROMName) -eq $true) {
        $strNewROMName = $hashtableAllMappings.Item($strThisLocalDATROMName)
        ($arrMachineProperties | Where-Object { $_.Name -eq $strLocalDATToCompareToMAMEPrimaryKeyColumnHeader }).Value = $strNewROMName
    } else {
        Write-Warning ('Somehow, our ROM mapping hashtable did not contain a key (in a key-value pair) for ' + $strThisLocalDATROMName + '. So we will not adjust its "ROMName" (primary key) in the revised DAT file. This may result in a naming collision')
    }
    $objThisMachineToCompareToMAME
    $intCounter++
}

Write-Verbose ('Exporting matching results to CSV: ' + $strCSVOutputFileRenamedROMDAT)
$arrRevisedDAT | Sort-Object -Property $strLocalDATToCompareToMAMEPrimaryKeyColumnHeader |
    Export-Csv -Path $strCSVOutputFileRenamedROMDAT -NoTypeInformation

Write-Verbose "Done!"

$VerbosePreference = $actionPreferenceFormerVerbose
$DebugPreference = $actionPreferenceFormerDebug
