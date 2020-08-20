# Convert-ProgettoEmmaCatverIniToCsv.ps1 is designed to take the English catver.ini file from
# http://www.progettoemma.net/history/catlist.php and convert it into tabular format in a CSV.
# In doing so, the category information and MAME version information (i.e., the first version
# of MAME that supports the ROM in some capacity) for each ROM can be combined with other data
# sources (e.g., using Join-Object in PowerShell, Power BI, SQL Server, or another tool of
# choice) to make a ROM list.

$strThisScriptVersionNumber = [version]'1.0.20200820.0'

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

#region Inputs
###############################################################################################
# Download catver.ini file from http://www.progettoemma.net/history/catlist.php and put it in
# the following folder:
# .\Progetto_Emma_Resources
# or if on Linux / MacOS: ./Progetto_Emma_Resources
# i.e., the folder that this script is in should have a subfolder called:
# Progetto_Emma_Resources
$strSubfolderPath = Join-Path '.' 'Progetto_Emma_Resources'

# The file will be processed and output as a CSV to
# .\Progetto_Emma_Category_and_MAME_Version_Information.csv
# or if on Linux / MacOS: ./Progetto_Emma_Category_and_MAME_Version_Information.csv
$strCSVOutputFile = Join-Path '.' 'Progetto_Emma_Category_and_MAME_Version_Information.csv'

# Display verbose output
$actionPreferenceFormerVerbose = $VerbosePreference
$VerbosePreference = [System.Management.Automation.ActionPreference]::Continue
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

function Convert-OneSelectedHashTableOfAttributes {
    # This function reads a hashtable from a hashtable of hashtables, then converts it into a
    # tabular data set. This function is designed to work with various MAME UI programs' .ini
    # files that have been converted to hashtables using Convert-IniToHashTable. This function
    # appends its output to the hashtable specified in the first argument. The output is a
    # hashtable (key-value pair) in the form of:
    # Key: primary key of tabular data
    # Value: PSCustomObject representing collected tabular data
    #
    # Twelve positional arguments are required:
    #
    # The first argument is a reference to an object that will be used to store output
    # The second argument is a reference to an object that serves as input. It is a "hashtable
    #   of hashtables" resulting, resulting from the collection of data using
    #   Convert-IniToHashTable
    # The third argument is a string representing the key of the input's outer hashtable. It
    #   "selects" the innner hashtable.
    # The fourth argument is either set to $null, or it's a string. If it's a string, it can
    #   either be an empty string ('') or it can be the name of one of the inner hashtable's
    #   keys, used to select the key for processing. If set to $null or '', the function
    #   assumes all inner hashtable keys need to be processed unless specified otherwise in
    #   argument five. If not set to $null or '', the function processes just the inner
    #   hashtable specified and ignores any others.
    # The fifth argument is only used if the fourth argument is not $null and not ''. In this
    #   case, it is a boolean. If set to $true, then the presence of an item in the selected
    #   hashtable is presumed to mean "affirmative" and the absense of an item is preseumed to
    #   mean "negative". See arguments 7 and 8. On the other hand, if the fifth argument is set
    #   to $false, then it is presumed that the items within the inner hashtable are key-value
    #   pairs (an inner-inner hashtable), and the key represents the item while the value
    #   represents the value for our tabular cell.
    # The sixth argument is a reference to an array. If the array has any elements, they are
    #   strings representing keys from the input's inner hashtable to ignore.
    # The seventh argument is the property name (column) to use in the output for storing the
    #   processed results
    # The eighth argument is an arbitrary object used as default, i.e., for the absense of an
    #   indicator. Usually this is 'False' or 'Unknown' - or similar.
    # The ninth argument is used only when the fourth argument is not $null or '' and the
    #   function is processing one key from the inner hashtable. The presence of an item on the
    #   inner hashtable indicates an "affirmative" - and whatever is specified in this eighth
    #   argument is stored. Usually this is 'True'. If the fourth arguement is $null or '',
    #   pass $null as the eighth argument.
    # The tenth argument is the name of the column used as the primary key.
    # The eleventh argument is a somewhat-redundant column that indicates that the primary key was
    #   processed as part of the current data set. Something like 'DataSetNamePresent' is
    #   appropriate.
    # The twelveth argument is a reference to an array of property names. Each time a new
    #   property is processed, its metadata is appended to the array and used for later calls
    #   to this function or for downstream post-processing.
    #
    # The function returns a 0 if successful, non-zero otherwise.
    #
    # Example usage #1 (Select one key from inner hashtable and treat as boolean):
    # $hashtableOutput = New-BackwardCompatibleCaseInsensitiveHashtable
    # $arrPropertyNamesAndDefaultValuesSoFar = @()
    # $strPropertyNameIndicatingDefinitionInHashTable = 'ProgettoSnapsCategoryPresent'
    # $strSubfolderPath = Join-Path '.' 'Progetto_Snaps_Resources'
    # $strFilePathProgettoSnapsCategoryArcadeIni = Join-Path $strSubfolderPath 'arcade.ini'
    # $strPropertyName = 'ProgettoSnapsCategoryArcade'
    # $objDefaultValue = 'False'
    # $strSectionNameOfSingleSectionToProcess = 'ROOT_FOLDER'
    # $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtablePrimary) $strFilePathProgettoSnapsCategoryArcadeIni $strSectionNameOfSingleSectionToProcess $true ([ref]($null)) $strPropertyName $objDefaultValue 'True' 'ROM' $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)
    #
    # Example usage #2 (Select one key from inner hashtable and process key-value pair (value
    #   is value for cell in tabular model)):
    # $hashtableOutput = New-BackwardCompatibleCaseInsensitiveHashtable
    # $arrPropertyNamesAndDefaultValuesSoFar = @()
    # $strPropertyNameIndicatingDefinitionInHashTable = 'ProgettoSnapsCategoryPresent'
    # $strSubfolderPath = Join-Path '.' 'Progetto_Snaps_Resources'
    # $strFilePathProgettoSnapsCategoryArcadeIni = Join-Path $strSubfolderPath 'arcade.ini'
    # $strPropertyName = 'ProgettoSnapsCategoryArcade'
    # $objDefaultValue = 'Unknown'
    # $strSectionNameOfSingleSectionToProcess = 'ROOT_FOLDER'
    # $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtablePrimary) $strFilePathProgettoSnapsCategoryArcadeIni $strSectionNameOfSingleSectionToProcess $true ([ref]($null)) $strPropertyName $objDefaultValue $null 'ROM' $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)
    #
    # Example usage #3 (Process all keys from inner hashtable with a few exceptions):
    # $hashtableOutput = New-BackwardCompatibleCaseInsensitiveHashtable
    # $arrPropertyNamesAndDefaultValuesSoFar = @()
    # $strPropertyNameIndicatingDefinitionInHashTable = 'ProgettoSnapsCategoryPresent'
    # $strSubfolderPath = Join-Path '.' 'Progetto_Snaps_Resources'
    # $strFilePathProgettoSnapsCategoryCabinetsIni = Join-Path $strSubfolderPath 'cabinets.ini'
    # $strPropertyName = 'ProgettoSnapsCategoryCabinetType'
    # $objDefaultValue = 'Unknown'
    # $arrIgnoreSections = @('FOLDER_SETTINGS', 'ROOT_FOLDER')
    # $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtablePrimary) $strFilePathProgettoSnapsCategoryCabinetsIni $null $null ([ref]$arrIgnoreSections) $strPropertyName 'Unknown' $null 'ROM' $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    $refHashtableOutput = $args[0]
    $refHashtableOfInputHashtables = $args[1]
    $strKeyToSelectInnerHashTable = $args[2] # $strFilePathProgettoSnapsCategoryArcadeIni
    $strSectionNameOfSingleSectionToProcess = $args[3] # 'ROOT_FOLDER'
    $boolTreatSingleSectionAsBoolean = $args[4]
    $refArrIgnoreSections = $args[5] # @('FOLDER_SETTINGS', 'ROOT_FOLDER')
    $strPropertyName = $args[6] # 'ProgettoSnapsCategoryArcade'
    $objDefaultValueForAbsenseOfIndicator = $args[7] # 'False'
    $objAffirmativeValueForPresenceOfIndicator = $args[8] # 'True'
    $strPrimaryKeyPropertyName = $args[9] # 'ROM'
    $strPropertyNameIndicatingDefinitionInHashTable = $args[10] # 'ProgettoSnapsCategoryPresent'
    $refArrPropertyNamesAndDefaultValuesSoFar = $args[11]

    $strThisFunctionVersionNumber = [version]'1.0.20200820.0'

    $intReturnCode = 0

    $boolMultivalued = $true
    if ($null -ne $strSectionNameOfSingleSectionToProcess) {
        if ('' -ne $strSectionNameOfSingleSectionToProcess) {
            $boolMultivalued = $false
        }
    }

    if (($refHashtableOfInputHashtables.Value).ContainsKey($strKeyToSelectInnerHashTable)) {
        if ($boolMultivalued -eq $false) {
            ($refHashtableOutput.Value).Keys | `
                ForEach-Object {
                    $strThisKey = $_
                    ($refHashtableOutput.Value).Item($strThisKey) | Add-Member -MemberType NoteProperty -Name $strPropertyName -Value $objDefaultValueForAbsenseOfIndicator
                }

            if (($refHashtableOfInputHashtables.Value).Item($strKeyToSelectInnerHashTable).ContainsKey($strSectionNameOfSingleSectionToProcess)) {
                ($refHashtableOfInputHashtables.Value).Item($strKeyToSelectInnerHashTable).Item($strSectionNameOfSingleSectionToProcess).Keys | `
                    ForEach-Object {
                        $strThisKey = $_
                        if (($refHashtableOutput.Value).ContainsKey($strThisKey)) {
                            if ($boolTreatSingleSectionAsBoolean) {
                                (($refHashtableOutput.Value).Item($strThisKey)).$strPropertyName = $objAffirmativeValueForPresenceOfIndicator
                            } else {
                                (($refHashtableOutput.Value).Item($strThisKey)).$strPropertyName = ($refHashtableOfInputHashtables.Value).Item($strKeyToSelectInnerHashTable).Item($strSectionNameOfSingleSectionToProcess).Item($strThisKey)
                            }
                        } else {
                            $PSCustomObjectROMMetadata = New-Object PSCustomObject
                            $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strPrimaryKeyPropertyName -Value $strThisKey
                            $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strPropertyNameIndicatingDefinitionInHashTable -Value 'True'
                            ($refArrPropertyNamesAndDefaultValuesSoFar.Value) | `
                                ForEach-Object {
                                    $strThisPropertyName = $_.PropertyName
                                    $objThisPropertyDefaultValue = $_.DefaultValue
                                    if ($_.MultivaluedProperty) {
                                        $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strThisPropertyName -Value @($objThisPropertyDefaultValue)
                                    } else {
                                        $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strThisPropertyName -Value $objThisPropertyDefaultValue
                                    }
                                }
                            if ($boolTreatSingleSectionAsBoolean) {
                                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strPropertyName -Value $objAffirmativeValueForPresenceOfIndicator
                            } else {
                                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strPropertyName -Value (($refHashtableOfInputHashtables.Value).Item($strKeyToSelectInnerHashTable).Item($strSectionNameOfSingleSectionToProcess).Item($strThisKey))
                            }
                            ($refHashtableOutput.Value).Add($strThisKey, $PSCustomObjectROMMetadata)
                        }
                    }
                $PSCustomObjectThisProperty = New-Object PSCustomObject
                $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'PropertyName' -Value $strPropertyName
                $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'DefaultValue' -Value $objDefaultValueForAbsenseOfIndicator
                $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'MultivaluedProperty' -Value $false
                ($refArrPropertyNamesAndDefaultValuesSoFar.Value) = ($refArrPropertyNamesAndDefaultValuesSoFar.Value) + $PSCustomObjectThisProperty
            } else {
                # Write-Error ('The following file had an unexpected file format and cannot be processed: ' + $strKeyToSelectInnerHashTable)
                $intReturnCode = 2
            }
        } else {
            $hashtableOutput.Keys | `
                ForEach-Object {
                    $strThisROMName = $_
                    $hashtableOutput.Item($strThisROMName) | Add-Member -MemberType NoteProperty -Name $strPropertyName -Value @($objDefaultValueForAbsenseOfIndicator)
                }

            ($refHashtableOfInputHashtables.Value).Item($strKeyToSelectInnerHashTable).Keys | `
                Where-Object { ($refArrIgnoreSections.Value) -notcontains $_ } | `
                Sort-Object | `
                ForEach-Object {
                    $strHeader = $_
                    (($refHashtableOfInputHashtables.Value).Item($strKeyToSelectInnerHashTable)).Item($strHeader).Keys | `
                        ForEach-Object {
                            $strThisKey = $_
                            if (($refHashtableOutput.Value).ContainsKey($strThisKey)) {
                                # ROM already on our output list
                                if (((($refHashtableOutput.Value).Item($strThisKey)).$strPropertyName).Count -eq 1) {
                                    # This multivalued attribute had one value stored
                                    if (((($refHashtableOutput.Value).Item($strThisKey)).$strPropertyName)[0] -eq $objDefaultValueForAbsenseOfIndicator) {
                                        # The existing value was the default value; replace it
                                        (($refHashtableOutput.Value).Item($strThisKey)).$strPropertyName = @($strHeader)
                                    } else {
                                        # The existing value was not the default value; append it so that we now have two values.
                                        (($refHashtableOutput.Value).Item($strThisKey)).$strPropertyName = (($refHashtableOutput.Value).Item($strThisKey)).$strPropertyName + $strHeader
                                    }
                                } else {
                                    # This multivalued attribute had more than one value stored; append this one
                                    (($refHashtableOutput.Value).Item($strThisKey)).$strPropertyName = (($refHashtableOutput.Value).Item($strThisKey)).$strPropertyName + $strHeader
                                }
                            } else {
                                # ROM was not on our output list
                                $PSCustomObjectROMMetadata = New-Object PSCustomObject
                                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strPrimaryKeyPropertyName -Value $strThisKey
                                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strPropertyNameIndicatingDefinitionInHashTable -Value 'True'
                                ($refArrPropertyNamesAndDefaultValuesSoFar.Value) | `
                                    ForEach-Object {
                                        $strThisPropertyName = $_.PropertyName
                                        $objThisPropertyDefaultValue = $_.DefaultValue
                                        if ($_.MultivaluedProperty) {
                                            $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strThisPropertyName -Value @($objThisPropertyDefaultValue)
                                        } else {
                                            $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strThisPropertyName -Value $objThisPropertyDefaultValue
                                        }
                                    }
                                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strPropertyName -Value @($strHeader)
                                ($refHashtableOutput.Value).Add($strThisKey, $PSCustomObjectROMMetadata)
                            }
                        }
                }
            $PSCustomObjectThisProperty = New-Object PSCustomObject
            $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'PropertyName' -Value $strPropertyName
            $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'DefaultValue' -Value $objDefaultValueForAbsenseOfIndicator
            $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'MultivaluedProperty' -Value $true
            ($refArrPropertyNamesAndDefaultValuesSoFar.Value) = ($refArrPropertyNamesAndDefaultValuesSoFar.Value) + $PSCustomObjectThisProperty
        }
    } else {
        $intReturnCode = 1
        # Write-Error ('Cannot process ROM information from the following file because it is missing in the hashtable: ' + $strKeyToSelectInnerHashTable)
    }

    $intReturnCode
}

$boolErrorOccurred = $false

# Progetto Emma catver.ini file
$strURLProgettoEmmaCatver = 'http://www.progettoemma.net/history/catlist.php'
$strFilePathProgettoEmmaCatverIni = Join-Path $strSubfolderPath 'catver.ini'

if ((Test-Path $strFilePathProgettoEmmaCatverIni) -ne $true) {
    Write-Error ('The Progetto Emma catver.ini file is missing. Please download it from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strURLProgettoEmmaCatver + "`n`n" + 'File Location:' + "`n" + $strFilePathProgettoEmmaCatverIni)
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

    $strFilePath = $strFilePathProgettoEmmaCatverIni
    $hashtableIniFile = $null
    Write-Verbose ('Ingesting data from file ' + $strFilePath + '...')
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtablePrimary.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ('An error occurred while procesing file ' + $strFilePath + ' and it will be skipped.')
    }

    ###########################################################################################

    # All files have been loaded into memory as hashtables at this point. Start transforming
    # data to form output.
    $hashtableOutput = New-BackwardCompatibleCaseInsensitiveHashtable
    $arrPropertyNamesAndDefaultValuesSoFar = @()
    $strPropertyNameIndicatingDefinitionInHashTable = 'ProgettoEmmaCategoryAndVersionInformationList'

    ###########################################################################################

    $strFilePath = $strFilePathProgettoEmmaCatverIni
    $strPropertyName = 'ProgettoEmmaCategoryAndVersionInformationCategory'
    $objDefaultValue = 'Unknown'
    $strSectionNameOfSingleSectionToProcess = 'Category'

    Write-Verbose ('Processing category data from file ' + $strFilePath + '...')
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtablePrimary) $strFilePath $strSectionNameOfSingleSectionToProcess $false ([ref]($null)) $strPropertyName $objDefaultValue $null 'ROM' $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    if ($intReturnCode -ne 0) {
        Write-Error ('An error occurred while procesing file ' + $strFilePath + '.')
    }

    ###########################################################################################

    $strPropertyNameParentCategory = 'ProgettoEmmaCategoryAndVersionInformationCategory'
    $strPropertyNameChildCategory = 'ProgettoEmmaCategoryAndVersionInformationSubcategory'
    $strPropertyNameMatureFlag = 'ProgettoEmmaCategoryAndVersionInformationMature'
    $strMatureSearchString = ' * Mature *'
    $objDefaultValue = 'Unknown'
    Write-Verbose ('Performing post-processing on category data from file ' + $strFilePath + '...')
    $hashtableOutput.Keys | `
        ForEach-Object {
            $strThisKey = $_
            $strThisFormerCombinedCategory = $hashtableOutput.Item($strThisKey).$strPropertyNameParentCategory
            
            if ($strThisFormerCombinedCategory.Contains($strMatureSearchString)) {
                $objMatureValue = 'True'
                $arrWorkingValue = Split-StringOnLiteralString $strThisFormerCombinedCategory $strMatureSearchString
            } else {
                $objMatureValue = 'False'
                $arrWorkingValue = @($strThisFormerCombinedCategory)
            }

            $arrCategories = Split-StringOnLiteralString ($arrWorkingValue[0]) ' / '
            $strParentCategory = $arrCategories[0]
            if ($arrCategories.Count -ge 2) {
                $strChildCategory = $arrCategories[1]
            } else {
                $strChildCategory = ''
            }

            $hashtableOutput.Item($strThisKey).$strPropertyNameParentCategory = $strParentCategory
            ($hashtableOutput.Item($strThisKey)) | Add-Member -MemberType NoteProperty -Name $strPropertyNameChildCategory -Value $strChildCategory
            ($hashtableOutput.Item($strThisKey)) | Add-Member -MemberType NoteProperty -Name $strPropertyNameMatureFlag -Value $objMatureValue
        }

    $PSCustomObjectThisProperty = New-Object PSCustomObject
    $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'PropertyName' -Value $strPropertyNameChildCategory
    $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'DefaultValue' -Value $objDefaultValue
    $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'MultivaluedProperty' -Value $false
    $arrPropertyNamesAndDefaultValuesSoFar = $arrPropertyNamesAndDefaultValuesSoFar + $PSCustomObjectThisProperty

    $PSCustomObjectThisProperty = New-Object PSCustomObject
    $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'PropertyName' -Value $strPropertyNameMatureFlag
    $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'DefaultValue' -Value $objDefaultValue
    $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name 'MultivaluedProperty' -Value $false
    $arrPropertyNamesAndDefaultValuesSoFar = $arrPropertyNamesAndDefaultValuesSoFar + $PSCustomObjectThisProperty

    ###########################################################################################

    $strFilePath = $strFilePathProgettoEmmaCatverIni
    $strPropertyName = 'ProgettoEmmaCategoryAndVersionInformationRomAddedToMameVersion'
    $objDefaultValue = 'Unknown'
    $strSectionNameOfSingleSectionToProcess = 'VerAdded'

    Write-Verbose ('Processing MAME version data from file ' + $strFilePath + '...')
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtablePrimary) $strFilePath $strSectionNameOfSingleSectionToProcess $false ([ref]($null)) $strPropertyName $objDefaultValue $null 'ROM' $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    if ($intReturnCode -ne 0) {
        Write-Error ('An error occurred while procesing file ' + $strFilePath + '.')
    }

    ###########################################################################################
    # All data has been tabularized; next, let's join the multivalued attributes' arrays
    Write-Verbose 'Performing Post-Processing...'

    $strJoining = ';'

    $arrJustMultiValuedAttributes = @($arrPropertyNamesAndDefaultValuesSoFar | `
        Where-Object {$_.MultivaluedProperty -eq $true} | `
        ForEach-Object {$_.PropertyName})

    if ($arrJustMultiValuedAttributes.Count -gt 0) {
        $hashtableOutput.Keys | `
            ForEach-Object {
                $strThisKey = $_
                $arrJustMultiValuedAttributes | `
                    ForEach-Object {
                        $strThisMultivaluedProperty = $_
                        $hashtableOutput.Item($strThisKey).$strThisMultivaluedProperty = $hashtableOutput.Item($strThisKey).$strThisMultivaluedProperty -join $strJoining
                    }
            }
    }

    # Write output file
    Write-Verbose 'Writing Output File...'
    $hashtableOutput.Values | Sort-Object -Property 'ROM' | Export-Csv $strCSVOutputFile -NoTypeInformation
    Write-Verbose 'Done'
}

$VerbosePreference = $actionPreferenceFormerVerbose
