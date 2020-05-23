# Convert-ProgettoSnapsBestGamesIniToCsv.ps1 is designed to take each of the "bestgames.ini"
# file from AntoPisa's website progettosnaps.net and convert it to a tabular CSV format. In
# doing so, the "quality score" that AntoPisa has assigned each game can be combined with other
# data sources (e.g., using Join-Object in PowerShell, Power BI, SQL Server, or another tool of
# choice) to make a ROM list.

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

# Download the bestgames.ini file from http://www.progettosnaps.net/bestgames/ and put it in
# the following folder:
# .\Progetto_Snaps_Resources
# or if on Linux / MacOS: ./Progetto_Snaps_Resources
# i.e., the folder that this script is in should have a subfolder called:
# Progetto_Snaps_Resources
$strSubfolderPath = Join-Path "." "Progetto_Snaps_Resources"

# The file will be processed and output as a CSV to
# .\Progetto_Snaps_Quality_Scores.csv
# or if on Linux / MacOS: ./Progetto_Snaps_Quality_Scores.csv
$strCSVOutputFile = Join-Path "." "Progetto_Snaps_Quality_Scores.csv"

###############################################################################################

function Split-StringOnLiteralString {
    # This function takes two positional arguments
    # The first argument is a string, and the string to be split
    # The second argument is a string or char, and it is that which is to split the string in the first parameter
    #
    # Wrap this function call in "a cast to array" to ensure that it always returns an array even when the result is a single string.
    # Example:
    # $result = @(Split-StringOnLiteralString "foo" " ")
    # # $result.GetType().FullName is System.Object[]
    # # $result.Count is 1

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

$boolErrorOccurred = $false

# Progetto Snaps "Best Games" ini file
$strURLProgettoSnapsBestGames = "http://www.progettosnaps.net/bestgames/"
$strFilePathProgettoSnapsBestGamesIni = Join-Path $strSubfolderPath "bestgames.ini"

if ((Test-Path $strFilePathProgettoSnapsBestGamesIni) -ne $true) {
    Write-Error ("The Progetto Snaps `"Best Games`" ini file is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLProgettoSnapsBestGames + "`n`nFile Location:`n" + $strFilePathProgettoSnapsBestGamesIni)
    $boolErrorOccurred = $true
}

if ($boolErrorOccurred -eq $false) {
    # We have all the files, let's do stuff

    $csvCurrentRomList = @()

    $strCurrentFilePath = $strFilePathProgettoSnapsBestGamesIni

    $arrStrFileContent = @(Get-Content $strCurrentFilePath)
    $intCurrentScore = $null
    $strCurrentScoreDescription = $null
    for ($intLineCounter = 0; $intLineCounter -lt $arrStrFileContent.Length; $intLineCounter++) {
        if (($arrStrFileContent[$intLineCounter]).Length -ge 1) {
            # There is data on this line (it's not just blank)

            $boolWasValidSectionHeaderLine = $false
            if (($arrStrFileContent[$intLineCounter]).Substring(0, 1) -eq "[") {
                # Possible start of a new ini section
                if (($arrStrFileContent[$intLineCounter]).Substring(($arrStrFileContent[$intLineCounter]).Length - 1, 1) -eq "]") {
                    # Line has both an opening square bracket and a closing square bracket; it's a new section.
                    $boolWasValidSectionHeaderLine = $true
                    # Question is: is it a section that we care about?
                    $strHeaderMinusSquareBraces = ($arrStrFileContent[$intLineCounter]).Substring(1, ($arrStrFileContent[$intLineCounter]).Length - 2)
                    $arrLineInProgress = @(Split-StringOnLiteralString ($strHeaderMinusSquareBraces.ToLower()) " to ")
                    if ($arrLineInProgress.Count -ge 2) {
                        # Header is in the format "x to y"
                        $intLowerScoreBoundary = [int]($arrLineInProgress[0])
                        $arrLineInProgress = @(Split-StringOnLiteralString ($arrLineInProgress[1]) " ")
                        if ($arrLineInProgress.Count -ge 2) {
                            $intUpperScoreBoundary = [int]($arrLineInProgress[0])
                            # Captured the upper and lower boundary; now, let's get the description
                            $arrLineInProgress = @(Split-StringOnLiteralString $strHeaderMinusSquareBraces "(")
                            $arrLineInProgress = @(Split-StringOnLiteralString ($arrLineInProgress[$arrLineInProgress.Count - 1]) ")")
                            if ($arrLineInProgress.Count -ge 2) {
                                $intCurrentScore = ($intLowerScoreBoundary + $intUpperScoreBoundary) / 2
                                $strCurrentScoreDescription = $arrLineInProgress[$arrLineInProgress.Count - 2]
                            }
                        }
                    }
                }
            }

            if ($boolWasValidSectionHeaderLine -eq $false) {
                if ($null -ne $intCurrentScore) {
                    # We are in a section that we care about and this line has data
                    # Let's assume it's a ROM
                    $strThisROMName = ($arrStrFileContent[$intLineCounter])
                    $result = @($csvCurrentRomList | Where-Object { $_.ROM -eq $strThisROMName })
                    if ($result.Count -ne 0) {
                        # ROM is already on the list
                        $csvCurrentRomList | Where-Object { $_.ROM -eq $strThisROMName } | `
                            ForEach-Object {
                                $_.ProgettoSnapsQualityList = "True"
                                if (($_.ProgettoSnapsQualityScore).Contains("`n" + ([string]$intCurrentScore) + "`n") -eq $false) {
                                    $_.ProgettoSnapsQualityScore = $_.ProgettoSnapsQualityScore + ([string]$intCurrentScore) + "`n"
                                }
                                if (($_.ProgettoSnapsQualityDescription).Contains("`n" + $strCurrentScoreDescription + "`n") -eq $false) {
                                    $_.ProgettoSnapsQualityDescription = $_.ProgettoSnapsQualityDescription + $strCurrentScoreDescription + "`n"
                                }
                            }
                    } else {
                        $PSCustomObjectROMMetadata = New-Object PSCustomObject
                        $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name "ROM" -Value $strThisROMName
                        $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name "ProgettoSnapsQualityList" -Value "True"
                        $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name "ProgettoSnapsQualityScore" -Value ("`n" + ([string]$intCurrentScore) + "`n")
                        $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name "ProgettoSnapsQualityDescription" -Value ("`n" + $strCurrentScoreDescription + "`n")
                        $csvCurrentRomList = $csvCurrentRomList + $PSCustomObjectROMMetadata
                    }
                }
            }
        }
    }

    # Clean up the tabular data
    $csvCurrentRomList = $csvCurrentRomList | `
        ForEach-Object {
            $doubleTotalScores = 0
            $intCountOfScores = 0
            $arrLineInProgress = Split-StringOnLiteralString ($_.ProgettoSnapsQualityScore) "`n"
            for ($intArrayCounter = 1; $intArrayCounter -le ($arrLineInProgress.Count - 2); $intArrayCounter++) {
                $doubleTotalScores = $doubleTotalScores + ([double]($arrLineInProgress[$intArrayCounter]))
                $intCountOfScores++
            }
            if ($intCountOfScores -ne 0) {
                $_.ProgettoSnapsQualityScore = [string]($doubleTotalScores / $intCountOfScores)
            }

            $strDescriptionLine = ""
            $arrLineInProgress = Split-StringOnLiteralString ($_.ProgettoSnapsQualityDescription) "`n"
            for ($intArrayCounter = 1; $intArrayCounter -le ($arrLineInProgress.Count - 2); $intArrayCounter++) {
                if ("" -eq $strDescriptionLine) {
                    $strDescriptionLine = $arrLineInProgress[$intArrayCounter]
                } else {
                    $strDescriptionLine = $strDescriptionLine + ";" + $arrLineInProgress[$intArrayCounter]
                }
            }
            if ("" -ne $strDescriptionLine) {
                $_.ProgettoSnapsQualityDescription = $strDescriptionLine
            }

            $_
        }

    $csvCurrentRomList | Export-Csv $strCSVOutputFile -NoTypeInformation
}
