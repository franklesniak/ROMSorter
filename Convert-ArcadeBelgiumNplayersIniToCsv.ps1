# Convert-ArcadeBelgiumNplayersIniToCsv.ps1 is designed to take the nplayers.ini file from
# http://nplayers.arcadebelgium.be/ and convert it into tabular format in a CSV. In doing so,
# the number of players for a particular game can be combined with other data sources (e.g.,
# using Join-Object in PowerShell, Power BI, SQL Server, or another tool of choice) to make a
# ROM list.
#
# This script takes a very long time to run, but I do not intend to run it often, and it works,
# so I am not planning to do anything about it right now :)

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

# Download nplayers.ini file from http://nplayers.arcadebelgium.be/ and put it in the following
# folder:
# .\Arcade_Belgium_Resources
# or if on Linux / MacOS: ./Arcade_Belgium_Resources
# i.e., the folder that this script is in should have a subfolder called:
# Arcade_Belgium_Resources
$strSubfolderPath = Join-Path "." "Arcade_Belgium_Resources"

# The file will be processed and output as a CSV to
# .\Arcade_Belgium_Number_Of_Players.csv
# or if on Linux / MacOS: ./Arcade_Belgium_Number_Of_Players.csv
$strCSVOutputFile = Join-Path "." "Arcade_Belgium_Number_Of_Players.csv"

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

$boolErrorOccurred = $false

# Arcade Belgium nplayers.ini file
$strURLArcadeBelgiumNplayers = "http://nplayers.arcadebelgium.be/"
$strFilePathArcadeBelgiumNplayersIni = Join-Path $strSubfolderPath "nplayers.ini"

if ((Test-Path $strFilePathArcadeBelgiumNplayersIni) -ne $true) {
    Write-Error ("The Arcade Belgium nplayers.ini file is missing. Please download it from the following URL and place it in the following location.`n`nURL: " + $strURLArcadeBelgiumNplayers + "`n`nFile Location:`n" + $strFilePathArcadeBelgiumNplayersIni)
    $boolErrorOccurred = $true
}

if ($boolErrorOccurred -eq $false) {
    # We have all the files, let's do stuff

    $csvCurrentRomList = @()

    $strCurrentFilePath = $strFilePathArcadeBelgiumNplayersIni

    $arrStrFileContent = @(Get-Content $strCurrentFilePath)
    $strHeaderMinusSquareBraces = $null
    for ($intLineCounter = 0; $intLineCounter -lt $arrStrFileContent.Length; $intLineCounter++) {
        if (($arrStrFileContent[$intLineCounter]).Length -ge 1) {
            # There is data on this line (it's not just blank)

            if ($intLineCounter -ge 10) {
                Write-Progress -Activity "Converting Arcade Belgium nplayers.ini file to CSV" -Status "Processing" -PercentComplete (($intLineCounter) / ($arrStrFileContent.Length) * 100)
            }
            $boolIsComment = $false
            if (($arrStrFileContent[$intLineCounter]).Length -ge 3) {
                if (($arrStrFileContent[$intLineCounter]).Substring(0, 3) -eq ";; ") {
                    if (($arrStrFileContent[$intLineCounter]).Substring(($arrStrFileContent[$intLineCounter]).Length - 3, 3) -eq " ;;") {
                        $boolIsComment = $true
                    }
                }
            }

            if ($boolIsComment -eq $false) {
                $boolWasValidSectionHeaderLine = $false
                if (($arrStrFileContent[$intLineCounter]).Substring(0, 1) -eq "[") {
                    # Possible start of a new ini section
                    if (($arrStrFileContent[$intLineCounter]).Substring(($arrStrFileContent[$intLineCounter]).Length - 1, 1) -eq "]") {
                        # Line has both an opening square bracket and a closing square bracket; it's a new section.
                        $boolWasValidSectionHeaderLine = $true
                        $strHeaderMinusSquareBraces = ($arrStrFileContent[$intLineCounter]).Substring(1, ($arrStrFileContent[$intLineCounter]).Length - 2)
                    }
                }

                if ($boolWasValidSectionHeaderLine -eq $false) {
                    if ($null -ne $strHeaderMinusSquareBraces) {
                        # We are in a section and this line has data
                        # Let's assume it's a ROM
                        $arrLineInProgress = Split-StringOnLiteralString ($arrStrFileContent[$intLineCounter]) "="
                        if ($arrLineInProgress.Count -eq 2) {
                            $strThisROMName = $arrLineInProgress[0]
                            $strNumberOfPlayersForThisROM = $arrLineInProgress[1]
                            $arrNumberOfPlayersForThisROM = Split-StringOnLiteralString $strNumberOfPlayersForThisROM " / "

                            $result = @($csvCurrentRomList | Where-Object { $_.ROM -eq $strThisROMName })
                            if ($result.Count -ne 0) {
                                # ROM is already on the list
                                for ($intCounterA = 0; $intCounterA -lt $result.Count; $intCounterA++) {
                                    if (($result[$intCounterA]).ROM -eq $strThisROMName) {
                                        ($result[$intCounterA]).ArcadeBelgiumNumberOfPlayersList = "True"
                                        if ((($result[$intCounterA]).ArcadeBelgiumNumberOfPlayers).Contains("`n" + ([string]::Join("`n", $arrNumberOfPlayersForThisROM)) + "`n") -eq $false) {
                                            ($result[$intCounterA]).ArcadeBelgiumNumberOfPlayers = ($result[$intCounterA]).ArcadeBelgiumNumberOfPlayers + ([string]::Join("`n", $arrNumberOfPlayersForThisROM)) + "`n"
                                        }
                                    }
                                }
                            } else {
                                $PSCustomObjectROMMetadata = New-Object PSCustomObject
                                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name "ROM" -Value $strThisROMName
                                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name "ArcadeBelgiumNumberOfPlayersList" -Value "True"
                                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name "ArcadeBelgiumNumberOfPlayers" -Value ("`n" + ([string]::Join("`n", $arrNumberOfPlayersForThisROM)) + "`n")
                                $csvCurrentRomList = $csvCurrentRomList + $PSCustomObjectROMMetadata
                            }
                        }
                    }
                }
            }
        }
    }

    # Clean up the tabular data
    $intLineCounter = 0
    $csvCurrentRomList = $csvCurrentRomList | `
        ForEach-Object {
            if ($intLineCounter -ge 10) {
                Write-Progress -Activity "Converting Arcade Belgium nplayers.ini file to CSV" -Status "Cleaning up" -PercentComplete (($intLineCounter) / ($csvCurrentRomList.Count) * 100)
            }
            $strNumberOfPlayersLine = ""
            $arrLineInProgress = Split-StringOnLiteralString ($_.ArcadeBelgiumNumberOfPlayers) "`n"
            for ($intArrayCounter = 1; $intArrayCounter -le ($arrLineInProgress.Count - 2); $intArrayCounter++) {
                if ("" -eq $strNumberOfPlayersLine) {
                    $strNumberOfPlayersLine = $arrLineInProgress[$intArrayCounter]
                } else {
                    $strNumberOfPlayersLine = $strNumberOfPlayersLine + ";" + $arrLineInProgress[$intArrayCounter]
                }
            }
            if ("" -ne $strNumberOfPlayersLine) {
                $_.ArcadeBelgiumNumberOfPlayers = $strNumberOfPlayersLine
            }

            $intLineCounter++

            $_
        }

    $csvCurrentRomList | Export-Csv $strCSVOutputFile -NoTypeInformation
}
