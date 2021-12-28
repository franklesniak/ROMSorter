# Convert-ProgettoSnapsRenameSetIniToCsv.ps1 converts the AntoPISA RenameSet to a tabular data
# format that is easier to ingest and use in downstream scripts.

$strThisScriptVersionNumber = [version]'1.0.20211228.0'

#region License
###############################################################################################
# Copyright 2021 Frank Lesniak

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

#region Inputs
###############################################################################################
# Download the renameSET.ini file from http://www.progettosnaps.net/renameset/ and put it in
# the following folder:
# .\Progetto_Snaps_Resources
# or if on Linux / MacOS: ./Progetto_Snaps_Resources
# i.e., the folder that this script is in should have a subfolder called:
# Progetto_Snaps_Resources
$strSubfolderPath = Join-Path '.' 'Progetto_Snaps_Resources'

# The file will be processed and output as a CSV to
# .\Progetto_Snaps_RenameSet.csv
# or if on Linux / MacOS: ./Progetto_Snaps_RenameSet.csv
$strCSVOutputFile = Join-Path '.' 'Progetto_Snaps_RenameSet.csv'

# Comment-out the following line if you prefer that the script operate silently.
$actionPreferenceNewVerbose = [System.Management.Automation.ActionPreference]::Continue
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

function Convert-IniToHashTable {
    # This function reads an .ini file and converts it to a hashtable
    #
    # Five or six positional arguments are required:
    #
    # The first argument is a reference to an object that will be used to store output
    # The second argument is a string representing the file path to the ini file
    # The third argument is an array of characters that represent the characters allowed to
    #   indicate the start of a comment. Usually, this should be set to @(';'), but if hashtags
    #   are also allowed as comments for a given application, then it should be set to
    #   @(';', '#') or @('#')
    # The fourth argument is a boolean value that indicates whether comments should be ignored.
    #   Normally, comments should be ignored, and so this should be set to $true
    # The fifth argument is a boolean value that indicates whether comments must be on their
    #   own line in order to be considered a comment. If set to $false, and if the semicolon
    #   is the character allowed to indicate the start of a comment, then the text after the
    #   semicolon in this example would not be considered a comment:
    #   key=value ; this text would not be considered a comment
    #   in this example, the value would be:
    #   value ; this text would not be considered a comment
    # The sixth argument is a string representation of the null section name. In other words,
    #   if a key-value pair is found outside of a section, what should be used as its fake
    #   section name? As an example, this can be set to 'NoSection' as long as their is no
    #   section in the ini file like [NoSection]
    # The seventh argument is a boolean value that indicates whether it is permitted for keys
    #   in the ini file to be supplied without an equal sign (if $true, the key is ingested but
    #   the value is regarded as $null). If set to false, lines that lack an equal sign are
    #   considered invalid and ignored.
    # If supplied, the eighth argument is a string representation of the comment prefix and is
    #   to being the name of the 'key' representing the comment (and appended with an index
    #   number beginning with 1). If argument four is set to $false, then this argument is
    #   required. Usually 'Comment' is OK to use, unless there are keys in the file named like
    #   'Comment1', 'Comment2', etc.
    #
    # The function returns a 0 if successful, non-zero otherwise.
    #
    # Example usage:
    # $hashtableConfigIni = $null
    # $intReturnCode = Convert-IniToHashTable ([ref]$hashtableConfigIni) '.\config.ini' @(';') $true $true 'NoSection' $true
    #
    # This function is derived from Get-IniContent at the website:
    # https://github.com/lipkau/PsIni/blob/master/PSIni/Functions/Get-IniContent.ps1
    # retrieved on 2020-05-30
    #region OriginalLicense
    # Although substantial modifications have been made, the original portions of
    # Get-IniContent that are incorporated into Convert-IniToHashTable are subject to the
    # following license:
    ###############################################################################################
    # Copyright 2019 Oliver Lipkau

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
    #endregion OriginalLicense

    $refOutput = $args[0]
    $strFilePath = $args[1]
    $arrCharCommentIndicator = $args[2]
    $boolIgnoreComments = $args[3]
    $boolCommentsMustBeOnOwnLine = $args[4]
    $strNullSectionName = $args[5]
    $boolAllowKeysWithoutValuesThatOmitEqualSign = $args[6]
    if ($boolIgnoreComments -ne $true) {
        $strCommentPrefix = $args[7]
    }

    $strThisFunctionVersionNumber = [version]'1.0.20200818.0'

    # Initialize regex matching patterns
    $arrCharCommentIndicator = $arrCharCommentIndicator | ForEach-Object {
        [regex]::Escape($_)
    }
    $strRegexComment = '^\s*([' + ($arrCharCommentIndicator -join '') + '].*)$'
    $strRegexCommentAnywhere = '\s*([' + ($arrCharCommentIndicator -join '') + '].*)$'
    $strRegexSection = '^\s*\[(.+)\]\s*$'
    $strRegexKey = '^\s*(.+?)\s*=\s*([''"]?)(.*)\2\s*$'

    $hashtableIni = New-BackwardCompatibleCaseInsensitiveHashtable

    if ((Test-Path $strFilePath) -eq $false) {
        Write-Error ('Could not process INI file; the specified file was not found: ' + $strFilePath)
        1 # return failure code
    } else {
        $intCommentCount = 0
        $strSection = $null
        switch -regex -file $strFilePath {
            $strRegexSection {
                $strSection = $Matches[1]
                if ($hashtableIni.ContainsKey($strSection) -eq $false) {
                    $hashtableIni.Add($strSection, (New-BackwardCompatibleCaseInsensitiveHashtable))
                }
                $intCommentCount = 0
                continue
            }

            $strRegexComment {
                if ($boolIgnoreComments -ne $true) {
                    if ($null -eq $strSection) {
                        $strEffectiveSection = $strNullSectionName
                        if ($hashtableIni.ContainsKey($strEffectiveSection) -eq $false) {
                            $hashtableIni.Add($strEffectiveSection, (New-BackwardCompatibleCaseInsensitiveHashtable))
                        }
                    } else {
                        $strEffectiveSection = $strSection
                    }
                    $intCommentCount++
                    if (($hashtableIni.Item($strEffectiveSection)).ContainsKey($strCommentPrefix + ([string]$intCommentCount))) {
                        Write-Warning ('File "' + $strFilePath + '", section "' + $strEffectiveSection + '" already unexpectedly contains a key "' + ($strCommentPrefix + ([string]$intCommentCount)) + '" with value "' + ($hashtableIni.Item($strEffectiveSection)).Item($strCommentPrefix + ([string]$intCommentCount)) + '". Key''s value will be changed to: "' + $Matches[1] + '"')
                        ($hashtableIni.Item($strEffectiveSection)).Item($strCommentPrefix + ([string]$intCommentCount)) = $Matches[1]
                    } else {
                        ($hashtableIni.Item($strEffectiveSection)).Add($strCommentPrefix + ([string]$intCommentCount), $Matches[1])
                    }
                }
                continue
            }

            default {
                $strLine = $_
                if ($null -eq $strSection) {
                    $strEffectiveSection = $strNullSectionName
                    if ($hashtableIni.ContainsKey($strEffectiveSection) -eq $false) {
                        $hashtableIni.Add($strEffectiveSection, (New-BackwardCompatibleCaseInsensitiveHashtable))
                    }
                } else {
                    $strEffectiveSection = $strSection
                }

                $strKey = $null
                $strValue = $null
                if ($boolCommentsMustBeOnOwnLine) {
                    $arrLine = @([regex]::Split($strLine, $strRegexKey))
                    if ($arrLine.Count -ge 4) {
                        # Key-Value Pair found
                        $strKey = $arrLine[1]
                        $strValue = $arrLine[3]
                    } else {
                        # No key-value pair found
                        if ($boolAllowKeysWithoutValuesThatOmitEqualSign) {
                            if (($null -ne $arrLine[0]) -and ($arrLine[0]) -ne '') {
                                $strKey = $arrLine[0]
                            }
                        }
                    }
                } else {
                    # Comments do not have to be on their own line
                    $arrLine = @([regex]::Split($strLine, $strRegexCommentAnywhere))
                    # $arrLine[0] is the line before any comments
                    $arrLineKeyValue = @([regex]::Split($arrLine[0], $strRegexKey))
                    if ($arrLineKeyValue.Count -ge 4) {
                        # Key-Value Pair found
                        $strKey = $arrLineKeyValue[1]
                        $strValue = $arrLineKeyValue[3]
                    } else {
                        # No key-value pair found
                        if ($boolAllowKeysWithoutValuesThatOmitEqualSign) {
                            if (($null -ne $arrLineKeyValue[0]) -and ($arrLineKeyValue[0]) -ne '') {
                                $strKey = $arrLineKeyValue[0]
                            }
                        }
                    }
                    # if $arrLine.Count -gt 1, $arrLine[1] is the comment portion of the line
                    if ($arrLine.Count -gt 1) {
                        if ($boolIgnoreComments -ne $true) {
                            $intCommentCount++
                            if (($hashtableIni.Item($strEffectiveSection)).ContainsKey($strCommentPrefix + ([string]$intCommentCount))) {
                                Write-Warning ('File "' + $strFilePath + '", section "' + $strEffectiveSection + '" already unexpectedly contains a key "' + ($strCommentPrefix + ([string]$intCommentCount)) + '" with value "' + ($hashtableIni.Item($strEffectiveSection)).Item($strCommentPrefix + ([string]$intCommentCount)) + '". Key''s value will be changed to: "' + $Matches[1] + '"')
                                ($hashtableIni.Item($strEffectiveSection)).Item($strCommentPrefix + ([string]$intCommentCount)) = $Matches[1]
                            } else {
                                ($hashtableIni.Item($strEffectiveSection)).Add($strCommentPrefix + ([string]$intCommentCount), $Matches[1])
                            }
                        }
                    }
                }

                if ($null -ne $strKey) {
                    if (($hashtableIni.Item($strEffectiveSection)).ContainsKey($strKey)) {
                        Write-Warning ('File "' + $strFilePath + '", section "' + $strEffectiveSection + '" already unexpectedly contains a key "' + $strKey + '" with value "' + ($hashtableIni.Item($strEffectiveSection)).Item($strKey) + '". Key''s value will be changed to: null')
                        ($hashtableIni.Item($strEffectiveSection)).Item($strKey) = $strValue
                    } else {
                        ($hashtableIni.Item($strEffectiveSection)).Add($strKey, $strValue)
                    }
                }
                continue
            }
        }
        $refOutput.Value = $hashtableIni
        0 # return success code
    }
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

$boolErrorOccurred = $false

# Progetto Snaps RenameSet renameSET.ini file
$strURLProgettoSnapsRenameSet = 'www.progettosnaps.net/renameset/'
$strFilePathProgettoSnapsRenameSetIni = Join-Path $strSubfolderPath 'renameSET.ini'

if ((Test-Path $strFilePathProgettoSnapsRenameSetIni) -ne $true) {
    Write-Error ('The Progetto Snaps RenameSet file "renameSET.ini" is missing. Please download it from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strURLProgettoSnapsRenameSet + "`n`n" + 'File Location:' + "`n" + $strFilePathProgettoSnapsRenameSetIni)
    $boolErrorOccurred = $true
}

if ($boolErrorOccurred -eq $false) {
    # We have all the files, let's do stuff

    $hashtablePrimary = New-BackwardCompatibleCaseInsensitiveHashtable

    $arrCharCommentIndicator = @(';')
    $boolIgnoreComments = $true
    $boolCommentsMustBeOnOwnLine = $false
    $strNullSectionName = 'NoSection'
    $boolAllowKeysWithoutValuesThatOmitEqualSign = $true

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsRenameSetIni
    $hashtableIniFile = $null
    Write-Verbose ('Ingesting data from file ' + $strFilePath + '...')
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtablePrimary.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ('An error occurred while procesing file ' + $strFilePath + ' and it will be skipped.')
    }

    ###########################################################################################

    $arrCSVRenameSetInfo = $hashtablePrimary.Item($strFilePathProgettoSnapsRenameSetIni).Keys | Sort-Object |
        ForEach-Object {
            $strMAMEVersion = $_
            if ($strMAMEVersion -ne $strNullSectionName) {
                $PSCustomObject = New-Object PSCustomObject
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAMEVersion' -Value $strMAMEVersion

                $versionPowerShellFriendly = Convert-MAMEVersionNumberToRepresentativePowerShellVersion $strMAMEVersion
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAMEVersionPowerShellFriendly' -Value $versionPowerShellFriendly

                $strMAMEDate = $hashtablePrimary.Item($strFilePathProgettoSnapsRenameSetIni).Item($strMAMEVersion).Item('Stat_01')
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAMEDate' -Value $strMAMEDate

                $PSCustomObject
            }
        } | Sort-Object -Property 'MAMEVersionPowerShellFriendly' | ForEach-Object {
            $PSCustomObjectHeader = $_
            $strMAMEVersion = $PSCustomObjectHeader.MAMEVersion
            $versionPowerShellFriendly = $PSCustomObjectHeader.MAMEVersionPowerShellFriendly
            $strMAMEDate = $PSCustomObjectHeader.MAMEDate
            $arrAllKeys = @($hashtablePrimary.Item($strFilePathProgettoSnapsRenameSetIni).Item($strMAMEVersion).Keys)
            $arrDelKeys = @($arrAllKeys | ForEach-Object {
                    $strMAMEVersionMetadataItem = $_
                    if ($strMAMEVersionMetadataItem.ToLower().Contains('del_')) {
                        $strMAMEVersionMetadataItem
                    }
                })
            $arrRenKeys = @($arrAllKeys | ForEach-Object {
                    $strMAMEVersionMetadataItem = $_
                    if ($strMAMEVersionMetadataItem.ToLower().Contains('ren_')) {
                        $strMAMEVersionMetadataItem
                    }
                })

            $arrDelROMPackages = @($arrDelKeys | ForEach-Object {
                    $strDelKey = $_
                    $hashtablePrimary.Item($strFilePathProgettoSnapsRenameSetIni).Item($strMAMEVersion).Item($strDelKey)
                })
            $arrRenROMPackages = @($arrRenKeys | ForEach-Object {
                    $strRenKey = $_
                    $strRenInfo = $hashtablePrimary.Item($strFilePathProgettoSnapsRenameSetIni).Item($strMAMEVersion).Item($strRenKey)
                    $arrRenInfoA = Split-StringOnLiteralString $strRenInfo '> '
                    $arrRenInfoB = Split-StringOnLiteralString $arrRenInfoA[$arrRenInfoA.Length - 1] ' '
                    $strNewName = ($arrRenInfoB[0]).ToLower()
                    $arrRenInfoC = Split-StringOnLiteralString $strRenInfo ' '
                    $strOldName = ($arrRenInfoC[0]).ToLower()

                    $PSCustomObject = New-Object PSCustomObject
                    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'OldROMPackageName' -Value $strOldName
                    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'NewROMPackageName' -Value $strNewName
                    $PSCustomObject
                })

            $arrDelROMPackages | ForEach-Object {
                $strOldName = $_
                $strNewName = ''
                $PSCustomObject = New-Object PSCustomObject
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAMEVersion' -Value $strMAMEVersion
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAMEVersionPowerShellFriendly' -Value $versionPowerShellFriendly
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAMEDate' -Value $strMAMEDate
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'Operation' -Value 'D'
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'OldROMPackageName' -Value $strOldName
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'NewROMPackageName' -Value $strNewName
                $PSCustomObject
            }

            $arrRenROMPackages | ForEach-Object {
                $PSCustomObjectRenameInfo = $_
                $strOldName = $PSCustomObjectRenameInfo.OldROMPackageName
                $strNewName = $PSCustomObjectRenameInfo.NewROMPackageName
                $PSCustomObject = New-Object PSCustomObject
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAMEVersion' -Value $strMAMEVersion
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAMEVersionPowerShellFriendly' -Value $versionPowerShellFriendly
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAMEDate' -Value $strMAMEDate
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'Operation' -Value 'R'
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'OldROMPackageName' -Value $strOldName
                $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'NewROMPackageName' -Value $strNewName
                $PSCustomObject
            }
        }
    $arrCSVRenameSetInfo |
        Sort-Object -Property @('MAMEVersionPowerShellFriendly', 'Operation', 'OldROMPackageName') |
        Export-Csv -Path $strCSVOutputFile -NoTypeInformation
}
