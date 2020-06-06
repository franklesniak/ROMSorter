# Convert-ProgettoSnapsCategoryIniFilesToCsv.ps1 is designed to take the "category.ini" files
# from AntoPisa's website progettosnaps.net and convert then to a tabular CSV format. In doing
# so, additional category information for each ROM (e.g., a list of arcade ROMs excluding BIOS
# files, a list of cloned arcade ROMs, a list of nonworking ROMs, etc.) can be combined with
# other data sources (e.g., using Join-Object in PowerShell, Power BI, SQL Server, or another
# tool of choice) to make a ROM list.

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

# Download the "category.ini" ZIP file from http://www.progettosnaps.net/renameset/ and put
# each of the .ini files from the ZIP file in the following folder:
# .\Progetto_Snaps_Resources
# or if on Linux / MacOS: ./Progetto_Snaps_Resources
# i.e., the folder that this script is in should have a subfolder called:
# Progetto_Snaps_Resources
$strSubfolderPath = Join-Path "." "Progetto_Snaps_Resources"

# The file will be processed and output as a CSV to
# .\Progetto_Snaps_Category_Info.csv
# or if on Linux / MacOS: ./Progetto_Snaps_Category_Info.csv
$strCSVOutputFile = Join-Path "." "Progetto_Snaps_Category_Info.csv"

# Display verbose output
$actionPreferenceFormerVerbose = $VerbosePreference
$VerbosePreference = [System.Management.Automation.ActionPreference]::Continue

###############################################################################################

function Split-StringOnLiteralString {
    # This function takes two positional arguments
    # The first argument is a string, and the string to be split
    # The second argument is a string or char, and it is that which is to split the string in the first parameter
    #
    # Note: This function always returns an array, even when there is zero or one element in it.
    #
    # Example:
    # $result = Split-StringOnLiteralString "foo" " "
    # # $result.GetType().FullName is System.Object[]
    # # $result.Count is 1
    #
    # Example 2:
    # $result = Split-StringOnLiteralString "What do you think of this function?" " "
    # # $result.Count is 7

    trap {
        Write-Error "An error occurred using the Split-StringOnLiteralString function. This was most likely caused by the arguments supplied not being strings"
    }

    if ($args.Length -ne 2) {
        Write-Error "Split-StringOnLiteralString was called without supplying two arguments. The first argument should be the string to be split, and the second should be the string or character on which to split the string."
    } else {
        if ($null -eq $args[0]) {
            # String to be split was $null; return an empty array. Leading comma ensures that
            # PowerShell cooperates and returns the array as desired (without collapsing it)
            , @()
        } elseif ($null -eq $args[1]) {
            # Splitter was $null; return string to be split within an array (of one element).
            # Leading comma ensures that PowerShell cooperates and returns the array as desired
            # (without collapsing it
            , ($args[0])
        } else {
            if (($args[0]).GetType().Name -ne "String") {
                Write-Warning "The first argument supplied to Split-StringOnLiteralString was not a string. It will be attempted to be converted to a string. To avoid this warning, cast arguments to a string before calling Split-StringOnLiteralString."
                $strToSplit = [string]$args[0]
            } else {
                $strToSplit = $args[0]
            }

            if ((($args[1]).GetType().Name -ne "String") -and (($args[1]).GetType().Name -ne "Char")) {
                Write-Warning "The second argument supplied to Split-StringOnLiteralString was not a string. It will be attempted to be converted to a string. To avoid this warning, cast arguments to a string before calling Split-StringOnLiteralString."
                $strSplitter = [string]$args[1]
            } elseif (($args[1]).GetType().Name -eq "Char") {
                $strSplitter = [string]$args[1]
            } else {
                $strSplitter = $args[1]
            }

            $strSplitterInRegEx = [regex]::Escape($strSplitter)

            # With the leading comma, force encapsulation into an array so that an array is
            # returned even when there is one element:
            , [regex]::Split($strToSplit, $strSplitterInRegEx)
        }
    }
}

function New-BackwardCompatibleCaseInsensitiveHashtable {
    # Usage:
    # $hashtable = New-BackwardCompatibleCaseInsensitiveHashtable
    $cultureDoNotCare = [System.Globalization.CultureInfo]::InvariantCulture
    $caseInsensitiveHashCodeProvider = New-Object -TypeName "System.Collections.CaseInsensitiveHashCodeProvider" -ArgumentList @($cultureDoNotCare)
    $caseInsensitiveComparer = New-Object -TypeName "System.Collections.CaseInsensitiveComparer" -ArgumentList @($cultureDoNotCare)
    New-Object -TypeName "System.Collections.Hashtable" -ArgumentList @($caseInsensitiveHashCodeProvider, $caseInsensitiveComparer)
}

function Convert-IniToHashTable {
    # This function reads an .ini file and converts it to a hashtable
    #
    # Five or six positional arguments are required:
    #
    # The first argument is a reference to an object that will be used to store output
    # The second argument is a string representing the file path to the ini file
    # The third argument is an array of characters that represent the characters allowed to
    #   indicate the start of a comment. Usually, this should be set to @(";"), but if hashtags
    #   are also allowed as comments for a given application, then it should be set to
    #   @(";", "#") or @("#")
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
    #   section name? As an example, this can be set to "NoSection" as long as their is no
    #   section in the ini file like [NoSection]
    # The seventh argument is a boolean value that indicates whether it is permitted for keys
    #   in the ini file to be supplied without an equal sign (if $true, the key is ingested but
    #   the value is regarded as $null). If set to false, lines that lack an equal sign are
    #   considered invalid and ignored.
    # If supplied, the eighth argument is a string representation of the comment prefix and is
    #   to being the name of the "key" representing the comment (and appended with an index
    #   number beginning with 1). If argument four is set to $false, then this argument is
    #   required. Usually "Comment" is OK to use, unless there are keys in the file named like
    #   "Comment1", "Comment2", etc.
    #
    # The function returns a 0 if successful, non-zero otherwise.
    #
    # Example usage:
    # $hashtableConfigIni = $null
    # $intReturnCode = Convert-IniToHashTable ([ref]$hashtableConfigIni) ".\config.ini" @(";") $true $true "NoSection" $true
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

    # Initialize regex matching patterns
    $arrCharCommentIndicator = $arrCharCommentIndicator | ForEach-Object {
        [regex]::Escape($_)
    }
    $strRegexComment = "^\s*([$($arrCharCommentIndicator -join '')].*)$"
    $strRegexCommentAnywhere = "\s*([$($arrCharCommentIndicator -join '')].*)$"
    $strRegexSection = "^\s*\[(.+)\]\s*$"
    $strRegexKey = "^\s*(.+?)\s*=\s*(['`"]?)(.*)\2\s*$"

    $hashtableIni = New-BackwardCompatibleCaseInsensitiveHashtable

    if ((Test-Path $strFilePath) -eq $false) {
        Write-Error ("Could not process INI file; the specified file was not found: " + $strFilePath)
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
                        Write-Warning ("File `"" + $strFilePath + "`", section `"" + $strEffectiveSection + "`" already unexpectedly contains a key `"" + ($strCommentPrefix + ([string]$intCommentCount)) + "`" with value `"" + ($hashtableIni.Item($strEffectiveSection)).Item($strCommentPrefix + ([string]$intCommentCount)) + "`". Key's value will be changed to: `"" + $Matches[1] + "`"")
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
                            if (($null -ne $arrLine[0]) -and ("" -ne $arrLine[0])) {
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
                            if (($null -ne $arrLineKeyValue[0]) -and ("" -ne $arrLineKeyValue[0])) {
                                $strKey = $arrLineKeyValue[0]
                            }
                        }
                    }
                    # if $arrLine.Count -gt 1, $arrLine[1] is the comment portion of the line
                    if ($arrLine.Count -gt 1) {
                        if ($boolIgnoreComments -ne $true) {
                            $intCommentCount++
                            if (($hashtableIni.Item($strEffectiveSection)).ContainsKey($strCommentPrefix + ([string]$intCommentCount))) {
                                Write-Warning ("File `"" + $strFilePath + "`", section `"" + $strEffectiveSection + "`" already unexpectedly contains a key `"" + ($strCommentPrefix + ([string]$intCommentCount)) + "`" with value `"" + ($hashtableIni.Item($strEffectiveSection)).Item($strCommentPrefix + ([string]$intCommentCount)) + "`". Key's value will be changed to: `"" + $Matches[1] + "`"")
                                ($hashtableIni.Item($strEffectiveSection)).Item($strCommentPrefix + ([string]$intCommentCount)) = $Matches[1]
                            } else {
                                ($hashtableIni.Item($strEffectiveSection)).Add($strCommentPrefix + ([string]$intCommentCount), $Matches[1])
                            }
                        }
                    }
                }
                
                if ($null -ne $strKey) {
                    if (($hashtableIni.Item($strEffectiveSection)).ContainsKey($strKey)) {
                        Write-Warning ("File `"" + $strFilePath + "`", section `"" + $strEffectiveSection + "`" already unexpectedly contains a key `"" + $strKey + "`" with value `"" + ($hashtableIni.Item($strEffectiveSection)).Item($strKey) + "`". Key's value will be changed to: null")
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
    # Eleven positional arguments are required:
    #
    # The first argument is a reference to an object that will be used to store output
    # The second argument is a reference to an object that serves as input. It is a "hashtable
    #   of hashtables" resulting, resulting from the collection of data using
    #   Convert-IniToHashTable
    # The third argument is a string representing the key of the input's outer hashtable. It
    #   "selects" the innner hashtable.
    # The fourth argument is either set to $null, or it's a string. If it's a string, it can
    #   either be an empty string ("") or it can be the name of one of the inner hashtable's
    #   keys, used to select the key for processing. If set to $null or "", the function
    #   assumes all inner hashtable keys need to be processed unless specified otherwise in
    #   argument five. If not set to $null or "", the function processes just the inner
    #   hashtable specified and ignores any others. The presence of an item in the selected
    #   hashtable is presumed to mean "affirmative" and the absense of an item is preseumed to
    #   mean "negative". See arguments 7 and 8.
    # The fifth argument is a reference to an array. If the array has any elements, they are
    #   strings representing keys from the input's inner hashtable to ignore.
    # The sixth argument is the property name (column) to use in the output for storing the
    #   processed results
    # The seventh argument is an arbitrary object used as default, i.e., for the absense of an
    #   indicator. Usually this is "False" or "Unknown" - or similar.
    # The eighth argument is used only when the fourth argument is not $null or "" and the
    #   function is processing one key from the inner hashtable. The presence of an item on the
    #   inner hashtable indicates an "affirmative" - and whatever is specified in this eighth
    #   argument is stored. Usually this is "True". If the fourth arguement is $null or "",
    #   pass $null as the eighth argument.
    # The ninth argument is the name of the column used as the primary key.
    # The tenth argument is a somewhat-redundant column that indicates that the primary key was
    #   processed as part of the current data set. Something like "DataSetNamePresent" is
    #   appropriate.
    # The eleventh argument is a reference to an array of property names. Each time a new
    #   property is processed, its metadata is appended to the array and used for later calls
    #   to this function or for downstream post-processing.
    #
    # The function returns a 0 if successful, non-zero otherwise.
    #
    # Example usage #1 (Select one key from inner hashtable and treat as boolean):
    # $hashtableOutput = New-BackwardCompatibleCaseInsensitiveHashtable
    # $arrPropertyNamesAndDefaultValuesSoFar = @()
    # $strPropertyNameIndicatingDefinitionInHashTable = "ProgettoSnapsCategoryPresent"
    # $strSubfolderPath = Join-Path "." "Progetto_Snaps_Resources"
    # $strFilePathProgettoSnapsCategoryArcadeIni = Join-Path $strSubfolderPath "arcade.ini"
    # $strPropertyName = "ProgettoSnapsCategoryArcade"
    # $objDefaultValue = "False"
    # $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"
    # $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePathProgettoSnapsCategoryArcadeIni $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)
    #
    # Example usage #2 (Process all keys from inner hashtable with a few exceptions):
    # $hashtableOutput = New-BackwardCompatibleCaseInsensitiveHashtable
    # $arrPropertyNamesAndDefaultValuesSoFar = @()
    # $strPropertyNameIndicatingDefinitionInHashTable = "ProgettoSnapsCategoryPresent"
    # $strSubfolderPath = Join-Path "." "Progetto_Snaps_Resources"
    # $strFilePathProgettoSnapsCategoryCabinetsIni = Join-Path $strSubfolderPath "cabinets.ini"
    # $strPropertyName = "ProgettoSnapsCategoryCabinetType"
    # $objDefaultValue = "Unknown"
    # $arrIgnoreSections = @("FOLDER_SETTINGS", "ROOT_FOLDER")
    # $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePathProgettoSnapsCategoryCabinetsIni $null ([ref]$arrIgnoreSections) $strPropertyName "Unknown" $null "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    $refHashtableOutput = $args[0]
    $refHashtableOfInputHashtables = $args[1]
    $strKeyToSelectInnerHashTable = $args[2] # $strFilePathProgettoSnapsCategoryArcadeIni
    $strSectionNameThatIndicatesBooleanTrue = $args[3] # "ROOT_FOLDER"
    $refArrIgnoreSections = $args[4] # @("FOLDER_SETTINGS", "ROOT_FOLDER")
    $strPropertyName = $args[5] # "ProgettoSnapsCategoryArcade"
    $objDefaultValueForAbsenseOfIndicator = $args[6] # "False"
    $objAffirmativeValueForPresenceOfIndicator = $args[7] # "True"
    $strPrimaryKeyPropertyName = $args[8] # "ROM"
    $strPropertyNameIndicatingDefinitionInHashTable = $args[9] # "ProgettoSnapsCategoryPresent"
    $refArrPropertyNamesAndDefaultValuesSoFar = $args[10]

    $intReturnCode = 0

    $boolMultivalued = $true
    if ($null -ne $strSectionNameThatIndicatesBooleanTrue) {
        if ("" -ne $strSectionNameThatIndicatesBooleanTrue) {
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
            
            if (($refHashtableOfInputHashtables.Value).Item($strKeyToSelectInnerHashTable).ContainsKey($strSectionNameThatIndicatesBooleanTrue)) {
                ($refHashtableOfInputHashtables.Value).Item($strKeyToSelectInnerHashTable).Item($strSectionNameThatIndicatesBooleanTrue).Keys | `
                    ForEach-Object {
                        $strThisKey = $_
                        if (($refHashtableOutput.Value).ContainsKey($strThisKey)) {
                            (($refHashtableOutput.Value).Item($strThisKey)).$strPropertyName = $objAffirmativeValueForPresenceOfIndicator
                        } else {
                            $PSCustomObjectROMMetadata = New-Object PSCustomObject
                            $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strPrimaryKeyPropertyName -Value $strThisKey
                            $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strPropertyNameIndicatingDefinitionInHashTable -Value "True"
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
                            $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strPropertyName -Value $objAffirmativeValueForPresenceOfIndicator
                            ($refHashtableOutput.Value).Add($strThisKey, $PSCustomObjectROMMetadata)
                        }
                    }
                $PSCustomObjectThisProperty = New-Object PSCustomObject
                $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name "PropertyName" -Value $strPropertyName
                $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name "DefaultValue" -Value $objDefaultValueForAbsenseOfIndicator
                $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name "MultivaluedProperty" -Value $false
                ($refArrPropertyNamesAndDefaultValuesSoFar.Value) = ($refArrPropertyNamesAndDefaultValuesSoFar.Value) + $PSCustomObjectThisProperty
            } else {
                # Write-Error ("The following file had an unexpected file format and cannot be processed: " + $strKeyToSelectInnerHashTable)
                $intReturnCode = 2
            }
        } else {
            $hashtableOutput.Keys | `
                ForEach-Object {
                    $strThisROMName = $_
                    $hashtableOutput.Item($strThisROMName) | Add-Member -MemberType NoteProperty -Name $strPropertyName -Value @($objDefaultValueForAbsenseOfIndicator)
                }
            
            ($refHashtableOfInputHashtables.Value).Item($strKeyToSelectInnerHashTable).Keys | `
                Where-Object {($refArrIgnoreSections.Value) -notcontains $_} | `
                Sort-Object | `
                ForEach-Object {
                    $strHeader = $_
                    (($refHashtableOfInputHashtables.Value).Item($strKeyToSelectInnerHashTable)).Item($strHeader).Keys | `
                        ForEach-Object {
                            $strThisKey =$_
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
                                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name $strPropertyNameIndicatingDefinitionInHashTable -Value "True"
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
            $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name "PropertyName" -Value $strPropertyName
            $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name "DefaultValue" -Value $objDefaultValueForAbsenseOfIndicator
            $PSCustomObjectThisProperty | Add-Member -MemberType NoteProperty -Name "MultivaluedProperty" -Value $true
            ($refArrPropertyNamesAndDefaultValuesSoFar.Value) = ($refArrPropertyNamesAndDefaultValuesSoFar.Value) + $PSCustomObjectThisProperty
        }
    } else {
        $intReturnCode = 1
        # Write-Error ("Cannot process ROM information from the following file because it is missing in the hashtable: " + $strKeyToSelectInnerHashTable)
    }

    $intReturnCode
}

$boolErrorOccurred = $false

# Progetto Snaps Category arcade.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryArcadeIni = Join-Path $strSubfolderPath "arcade.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryArcadeIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"arcade.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryArcadeIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category arcade_BIOS.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryArcadeBIOSIni = Join-Path $strSubfolderPath "arcade_BIOS.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryArcadeBIOSIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"arcade_BIOS.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryArcadeBIOSIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category arcade_NOBIOS.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryArcadeNOBIOSIni = Join-Path $strSubfolderPath "arcade_NOBIOS.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryArcadeNOBIOSIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"arcade_NOBIOS.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryArcadeNOBIOSIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category Bootlegs.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryBootlegsIni = Join-Path $strSubfolderPath "Bootlegs.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryBootlegsIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"Bootlegs.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryBootlegsIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category cabinets.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryCabinetsIni = Join-Path $strSubfolderPath "cabinets.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryCabinetsIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"cabinets.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryCabinetsIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category category.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryCategoryIni = Join-Path $strSubfolderPath "category.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryCategoryIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"category.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryCategoryIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category "Clones Arcade.ini" file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryClonesArcadeIni = Join-Path $strSubfolderPath "Clones Arcade.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryClonesArcadeIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"Clones Arcade.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryClonesArcadeIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category driver.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryDriverIni = Join-Path $strSubfolderPath "driver.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryDriverIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"driver.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryDriverIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category freeplay.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryFreeplayIni = Join-Path $strSubfolderPath "freeplay.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryFreeplayIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"freeplay.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryFreeplayIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category "Mechanical Arcade.ini" file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryMechanicalArcadeIni = Join-Path $strSubfolderPath "Mechanical Arcade.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryMechanicalArcadeIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"Mechanical Arcade.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryMechanicalArcadeIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category mess.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryMessIni = Join-Path $strSubfolderPath "mess.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryMessIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"mess.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryMessIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category monochrome.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryMonochromeIni = Join-Path $strSubfolderPath "monochrome.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryMonochromeIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"monochrome.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryMonochromeIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category "Non Bootlegs.ini" file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryNonBootlegsIni = Join-Path $strSubfolderPath "Non Bootlegs.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryNonBootlegsIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"Non Bootlegs.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryNonBootlegsIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category "Non Mechanical Arcade.ini" file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryNonMechanicalArcadeIni = Join-Path $strSubfolderPath "Non Mechanical Arcade.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryNonMechanicalArcadeIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"Non Mechanical Arcade.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryNonMechanicalArcadeIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category "Not Working Arcade.ini" file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryNotWorkingArcadeIni = Join-Path $strSubfolderPath "Not Working Arcade.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryNotWorkingArcadeIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"Not Working Arcade.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryNotWorkingArcadeIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category "Originals Arcade.ini" file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryOriginalsArcadeIni = Join-Path $strSubfolderPath "Originals Arcade.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryOriginalsArcadeIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"Originals Arcade.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryOriginalsArcadeIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category Players.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryPlayersIni = Join-Path $strSubfolderPath "Players.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryPlayersIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"Players.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryPlayersIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category resolution.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryResolutionIni = Join-Path $strSubfolderPath "resolution.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryResolutionIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"resolution.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryResolutionIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category screenless.ini file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryScreenlessIni = Join-Path $strSubfolderPath "screenless.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryScreenlessIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"screenless.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryScreenlessIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category "Working Arcade Clean.ini" file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryWorkingArcadeCleanIni = Join-Path $strSubfolderPath "Working Arcade Clean.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryWorkingArcadeCleanIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"Working Arcade Clean.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryWorkingArcadeCleanIni)
    $boolErrorOccurred = $true
}

# Progetto Snaps Category "Working Arcade.ini" file
$strURLProgettoSnapsCategory = "www.progettosnaps.net/renameset/"
$strFilePathProgettoSnapsCategoryWorkingArcadeIni = Join-Path $strSubfolderPath "Working Arcade.ini"

if ((Test-Path $strFilePathProgettoSnapsCategoryWorkingArcadeIni) -ne $true) {
    Write-Error ("The Progetto Snaps Category ini file `"Working Arcade.ini`" is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsCategory + "`n`nFile Location:`n" + $strFilePathProgettoSnapsCategoryWorkingArcadeIni)
    $boolErrorOccurred = $true
}

if ($boolErrorOccurred -eq $false) {
    # We have all the files, let's do stuff

    $hashtableMaster = New-BackwardCompatibleCaseInsensitiveHashtable

    $arrCharCommentIndicator = @(";")
    $boolIgnoreComments = $true
    $boolCommentsMustBeOnOwnLine = $false
    $strNullSectionName = "NoSection"
    $boolAllowKeysWithoutValuesThatOmitEqualSign = $true

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryArcadeIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryArcadeBIOSIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryArcadeNOBIOSIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryBootlegsIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryCabinetsIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryCategoryIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryClonesArcadeIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryDriverIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryFreeplayIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryMechanicalArcadeIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryMessIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryMonochromeIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryNonBootlegsIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryNonMechanicalArcadeIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryNotWorkingArcadeIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryOriginalsArcadeIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryPlayersIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryResolutionIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryScreenlessIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryWorkingArcadeCleanIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryWorkingArcadeIni
    $hashtableIniFile = $null
    Write-Verbose ("Ingesting data from file " + $strFilePath + "...")
    $intReturnCode = Convert-IniToHashTable ([ref]$hashtableIniFile) $strFilePath $arrCharCommentIndicator $boolIgnoreComments $boolCommentsMustBeOnOwnLine $strNullSectionName $boolAllowKeysWithoutValuesThatOmitEqualSign

    if ($intReturnCode -eq 0) {
        $hashtableMaster.Add($strFilePath, $hashtableIniFile)
    } else {
        Write-Error ("An error occurred while procesing file " + $strFilePath + " and it will be skipped.")
    }

    ###########################################################################################

    # All files have been loaded into memory as hashtables at this point. Start transforming
    # data to form output.
    $hashtableOutput = New-BackwardCompatibleCaseInsensitiveHashtable
    $arrPropertyNamesAndDefaultValuesSoFar = @()
    $strPropertyNameIndicatingDefinitionInHashTable = "ProgettoSnapsCategoryPresent"

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryArcadeIni
    $strPropertyName = "ProgettoSnapsCategoryArcade"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryArcadeBIOSIni
    $strPropertyName = "ProgettoSnapsCategoryBIOSOfArcade"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryArcadeNOBIOSIni
    $strPropertyName = "ProgettoSnapsCategoryArcadeExcludingBIOS"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryBootlegsIni
    $strPropertyName = "ProgettoSnapsCategoryBootleg"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryCabinetsIni
    $strPropertyName = "ProgettoSnapsCategoryCabinetType"
    $objDefaultValue = "Unknown"

    $arrIgnoreSections = @("FOLDER_SETTINGS", "ROOT_FOLDER")

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $null ([ref]$arrIgnoreSections) $strPropertyName "Unknown" $null "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryCategoryIni
    $strPropertyName = "ProgettoSnapsCategoryCategory"
    $objDefaultValue = "Unknown"

    $arrIgnoreSections = @("FOLDER_SETTINGS", "ROOT_FOLDER")

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $null ([ref]$arrIgnoreSections) $strPropertyName "Unknown" $null "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryClonesArcadeIni
    $strPropertyName = "ProgettoSnapsCategoryArcadeClone"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryDriverIni
    $strPropertyName = "ProgettoSnapsCategoryDriver"
    $objDefaultValue = "Unknown"

    $arrIgnoreSections = @("FOLDER_SETTINGS", "ROOT_FOLDER")

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $null ([ref]$arrIgnoreSections) $strPropertyName "Unknown" $null "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryFreeplayIni
    $strPropertyName = "ProgettoSnapsCategoryArcadeThatAllowsFreePlay"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "Free Play"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryMechanicalArcadeIni
    $strPropertyName = "ProgettoSnapsCategoryMechanicalArcade"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryMessIni
    $strPropertyName = "ProgettoSnapsCategoryNonArcadeViaMESS"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryMonochromeIni
    $strPropertyName = "ProgettoSnapsCategoryColorStatus"
    $objDefaultValue = "Color/Unknown/Not Applicable"

    $arrIgnoreSections = @("FOLDER_SETTINGS", "ROOT_FOLDER")

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $null ([ref]$arrIgnoreSections) $strPropertyName "Unknown" $null "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryNonBootlegsIni
    $strPropertyName = "ProgettoSnapsCategoryNotABootleg"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryNonMechanicalArcadeIni
    $strPropertyName = "ProgettoSnapsCategoryNotAMechanicalArcade"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryNotWorkingArcadeIni
    $strPropertyName = "ProgettoSnapsCategoryNonWorkingArcade"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryOriginalsArcadeIni
    $strPropertyName = "ProgettoSnapsCategoryArcadeParent"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryPlayersIni
    $strPropertyName = "ProgettoSnapsCategoryNumberOfPlayers"
    $objDefaultValue = "Unknown or N/A"

    $arrIgnoreSections = @("FOLDER_SETTINGS", "ROOT_FOLDER")

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $null ([ref]$arrIgnoreSections) $strPropertyName "Unknown" $null "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryResolutionIni
    $strPropertyName = "ProgettoSnapsCategoryResolution"
    $objDefaultValue = "Unknown or N/A"

    $arrIgnoreSections = @("FOLDER_SETTINGS", "ROOT_FOLDER")

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $null ([ref]$arrIgnoreSections) $strPropertyName "Unknown" $null "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryScreenlessIni
    $strPropertyName = "ProgettoSnapsCategoryDoesNotUseADisplay"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryWorkingArcadeCleanIni
    $strPropertyName = "ProgettoSnapsCategoryWorkingArcadeNotInAGenerallyUndesirableCategory"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################

    $strFilePath = $strFilePathProgettoSnapsCategoryWorkingArcadeIni
    $strPropertyName = "ProgettoSnapsCategoryWorkingArcade"
    $objDefaultValue = "False"
    $strSectionNameThatIndicatesBooleanTrue = "ROOT_FOLDER"

    Write-Verbose ("Processing data from file " + $strFilePath + "...")
    $intReturnCode = Convert-OneSelectedHashTableOfAttributes ([ref]$hashtableOutput) ([ref]$hashtableMaster) $strFilePath $strSectionNameThatIndicatesBooleanTrue ([ref]($null)) $strPropertyName $objDefaultValue "True" "ROM" $strPropertyNameIndicatingDefinitionInHashTable ([ref]$arrPropertyNamesAndDefaultValuesSoFar)

    ###########################################################################################
    # All data has been tabularized; next, let's join the multivalued attributes' arrays
    Write-Verbose "Performing Post-Processing..."

    $strJoining = ";"

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
    Write-Verbose "Writing Output File..."
    $hashtableOutput.Values | Sort-Object -Property "ROM" | Export-Csv $strCSVOutputFile -NoTypeInformation
    Write-Verbose "Done"
}

$VerbosePreference = $actionPreferenceFormerVerbose
