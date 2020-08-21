# Create-ConsolidatedAllKillerNoFillerGameList.ps1 is designed to take each of the "All Killer
# No Filler" game lists and merge them into one consolidated, tabular CSV output file with the
# name of the ROM as the primary key of the CSV/table. The resulting output file can be joined
# using Join-Object in PowerShell, Power BI, SQL Server, or another tool of choice to pull in
# additional ROM metadata.

$strThisScriptVersionNumber = [version]'1.0.20200821.0'

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
# Set target system
$strTargetSystem = 'Raspberry Pi'

# Download all of the "All Killer No Filler" batch files and place them in the following folder:
# .\All_Killer_No_Filler_Batch_Files
# or if on Linux / MacOS: ./All_Killer_No_Filler_Batch_Files
$strSubfolderPath = Join-Path '.' 'All_Killer_No_Filler_Batch_Files'

# The All Killer No Filler batch files are located at a few different URLs on
# http://forum.arcadecontrols.com/
# The exact URLs are embedded below in variables that begin with $strURL...
# Or, if you run this script and are missing the files, it will throw an error and tell you
# what's missing.

# The files will be processed, consolidated, and output as a CSV to
# .\All_Killer_No_Filler_Consolidated_List.csv
# or if on Linux / MacOS: ./All_Killer_No_Filler_Consolidated_List.csv
$strCSVOutputFile = Join-Path '.' 'All_Killer_No_Filler_Consolidated_List.csv'
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

function Merge-AllKillerNoFillerFile {
    # The first parameter is a reference to an array
    # The second parameter is a string representing the path to the All Killer No Filler batch file
    # The third parameter is a string representing the category, according to the All Killer No Filler batch file
    # The fourth parameter is a string representing the screen orientation, according to the All Killer No Filler batch file

    # Example: Merge-AllKillerNoFillerFile ([ref]$csvCurrentRomList) $strCurrentFilePath $strCurrentFileCategory $strCurrentFileScreenOrientation

    $refCsvCurrentRomList = $args[0]
    $strCurrentFilePath = $args[1]
    $strCurrentFileCategory = $args[2]
    $strCurrentFileScreenOrientation = $args[3]

    $arrStrFileContent = @(Get-Content $strCurrentFilePath)
    $arrStrRomList = @($arrStrFileContent | `
        ForEach-Object {
            if ($_.Length -ge 2) {
                if ($_.Substring(0, 2) -ne '::') {
                    $_ # Not commented-out -- send down pipeline
                }
            } else {
                $_
            }
        } | `
        ForEach-Object {
            if ($_.Length -ge 4) {
                if ($_.Substring(0, 4) -ne 'rem ') {
                    $_ # Not commented-out -- send down pipeline
                }
            } else {
                $_
            }
        } | `
        ForEach-Object {
            if ($_.Length -ge 3) {
                if ($_.Substring(0, 3) -ne 'md ') {
                    $_ # Not a "make directory" command -- send down pipeline
                }
            } else {
                $_
            }
        } | `
        ForEach-Object {
            if ($_.Length -ge 6) {
                if ($_.Substring(0, 6) -ne 'mkdir ') {
                    $_ # Not a "make directory" command -- send down pipeline
                }
            } else {
                $_
            }
        } | `
        ForEach-Object {
            if ($_.Length -ge 5) {
                if ($_.Substring(0, 5) -eq 'copy ') {
                    $_ # It's a copy command -- send down pipeline
                }
            }
        } | `
        ForEach-Object {
            if ($_.ToLower().Contains('.zip')) {
                $_ # Contains .zip string --- well-formatted line for us to process -- send down pipeline
            }
        } | `
        ForEach-Object {
            $arrTempResult = Split-StringOnLiteralString ($_.ToLower()) 'copy '
            if ($arrTempResult.Count -ge 2) {
                $arrTempResultTwo = Split-StringOnLiteralString ($arrTempResult[1]) '.zip'
                $arrTempResultTwo[0] # Return just the ROM name
            }
        })

    $arrStrRomList | `
        ForEach-Object {
            $strThisROMName = $_
            $result = @($refCsvCurrentRomList.Value | Where-Object { $_.ROM -eq $strThisROMName })
            if ($result.Count -ne 0) {
                # ROM is already on the list
                for ($intCounterA = 0; $intCounterA -lt ($refCsvCurrentRomList.Value).Count; $intCounterA++) {
                    if ((($refCsvCurrentRomList.Value)[$intCounterA]).ROM -eq $strThisROMName) {
                        (($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerList = 'True'
                        if (((($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerCategory).Contains($strCurrentFileCategory) -eq $false) {
                            (($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerCategory = (($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerCategory + ';' + $strCurrentFileCategory
                        }
                        if (((($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerScreenOrientation).Contains($strCurrentFileScreenOrientation) -eq $false) {
                            (($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerScreenOrientation = (($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerScreenOrientation + ';' + $strCurrentFileScreenOrientation
                        }
                    }
                }
            } else {
                $PSCustomObjectROMMetadata = New-Object PSCustomObject
                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name 'ROM' -Value $strThisROMName
                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name 'AllKillerNoFillerList' -Value 'True'
                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name 'AllKillerNoFillerCategory' -Value $strCurrentFileCategory
                $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name 'AllKillerNoFillerScreenOrientation' -Value $strCurrentFileScreenOrientation
                ($refCsvCurrentRomList.Value) = ($refCsvCurrentRomList.Value) + $PSCustomObjectROMMetadata
            }
        }
}

function Merge-ROMManuallyOntoAllKillerNoFillerList {
    # The first parameter is a reference to an array
    # The second parameter is a string representing the name of the ROM to merge onto the list manually (i.e., as an override)
    # The third parameter is a string representing the category, according to the All Killer No Filler batch file
    # The fourth parameter is a string representing the screen orientation, according to the All Killer No Filler batch file

    # Example: Merge-ROMManuallyOntoAllKillerNoFillerList ([ref]$csvCurrentRomList) $strThisROMName $strCurrentFileCategory $strCurrentFileScreenOrientation

    $refCsvCurrentRomList = $args[0]
    $strThisROMName = ($args[1]).ToLower()
    $strCurrentFileCategory = $args[2]
    $strCurrentFileScreenOrientation = $args[3]

    $result = @($refCsvCurrentRomList.Value | Where-Object { $_.ROM -eq $strThisROMName })
    if ($result.Count -ne 0) {
        # ROM is already on the list
        for ($intCounterA = 0; $intCounterA -lt ($refCsvCurrentRomList.Value).Count; $intCounterA++) {
            if ((($refCsvCurrentRomList.Value)[$intCounterA]).ROM -eq $strThisROMName) {
                (($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerList = 'True'
                if (((($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerCategory).Contains($strCurrentFileCategory) -eq $false) {
                    (($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerCategory = (($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerCategory + ';' + $strCurrentFileCategory
                }
                if (((($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerScreenOrientation).Contains($strCurrentFileScreenOrientation) -eq $false) {
                    (($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerScreenOrientation = (($refCsvCurrentRomList.Value)[$intCounterA]).AllKillerNoFillerScreenOrientation + ';' + $strCurrentFileScreenOrientation
                }
            }
        }
    } else {
        $PSCustomObjectROMMetadata = New-Object PSCustomObject
        $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name 'ROM' -Value $strThisROMName
        $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name 'AllKillerNoFillerList' -Value 'True'
        $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name 'AllKillerNoFillerCategory' -Value $strCurrentFileCategory
        $PSCustomObjectROMMetadata | Add-Member -MemberType NoteProperty -Name 'AllKillerNoFillerScreenOrientation' -Value $strCurrentFileScreenOrientation
        ($refCsvCurrentRomList.Value) = ($refCsvCurrentRomList.Value) + $PSCustomObjectROMMetadata
    }
}

$boolErrorOccurred = $false

# "All Killer, No Filler" Shoot-em-ups (SHMUPS) Game Lists
# First forum post has attachments in the form of batch files
$strURLAllKillerNoFillerSHMUPS = 'http://forum.arcadecontrols.com/index.php/topic,149578.msg1561125.html'
$strFilePathAllKillerNoFillerSHMUPSHorizontalBatch = Join-Path $strSubfolderPath '_horshmups.txt'
$strFilePathAllKillerNoFillerSHMUPSVerticalBatch = Join-Path $strSubfolderPath '_vertshmups.txt'

if (((Test-Path $strFilePathAllKillerNoFillerSHMUPSHorizontalBatch) -ne $true) -or ((Test-Path $strFilePathAllKillerNoFillerSHMUPSVerticalBatch) -ne $true)) {
    Write-Error ('The All Killer No Filler "SHMUPS" batch file(s) are missing. Please download them from the following URL and place them in the following locations.' + "`n`n" + 'URL: ' + $strURLAllKillerNoFillerSHMUPS + "`n`n" + 'File Locations:' + "`n" + $strFilePathAllKillerNoFillerSHMUPSHorizontalBatch + "`n" + $strFilePathAllKillerNoFillerSHMUPSVerticalBatch)
    $boolErrorOccurred = $true
}

# "All Killer, No Filler" "Versus" Fighters (VS Fighters) Game List
# First forum post has an attachment in the form of a batch file
$strURLAllKillerNoFillerVSFighters = 'http://forum.arcadecontrols.com/index.php/topic,149619.0.html'
$strFilePathAllKillerNoFillerVSFightersBatch = Join-Path $strSubfolderPath '_vsfighting.txt'

if ((Test-Path $strFilePathAllKillerNoFillerVSFightersBatch) -ne $true) {
    Write-Error ('The All Killer No Filler "VS Fighters" batch file is missing. Please download it from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strURLAllKillerNoFillerVSFighters + "`n`n" + 'File Location:' + "`n" + $strFilePathAllKillerNoFillerVSFightersBatch)
    $boolErrorOccurred = $true
}

# "All Killer, No Filler" Sports Game List
# First forum post has an attachment in the form of a batch file
$strURLAllKillerNoFillerSports = 'http://forum.arcadecontrols.com/index.php?topic=149640.0'
$strFilePathAllKillerNoFillerSportsBatch = Join-Path $strSubfolderPath '_sports.txt'

if ((Test-Path $strFilePathAllKillerNoFillerSportsBatch) -ne $true) {
    Write-Error ('The All Killer No Filler Sports batch file is missing. Please download it from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strURLAllKillerNoFillerSports + "`n`n" + 'File Location:' + "`n" + $strFilePathAllKillerNoFillerSportsBatch)
    $boolErrorOccurred = $true
}

# "All Killer, No Filler" Puzzle Game List
# First forum post has an attachment in the form of a batch file
$strURLAllKillerNoFillerPuzzle = 'http://forum.arcadecontrols.com/index.php?topic=149693.0'
$strFilePathAllKillerNoFillerPuzzleBatch = Join-Path $strSubfolderPath '_puzzle.txt'

if ((Test-Path $strFilePathAllKillerNoFillerPuzzleBatch) -ne $true) {
    Write-Error ('The All Killer No Filler Puzzle batch file is missing. Please download it from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strURLAllKillerNoFillerPuzzle + "`n`n" + 'File Location:' + "`n" + $strFilePathAllKillerNoFillerPuzzleBatch)
    $boolErrorOccurred = $true
}

# "All Killer, No Filler" Run 'n' Gun Game List
# First forum post has an attachment in the form of a batch file
$strURLAllKillerNoFillerRunNGun = 'http://forum.arcadecontrols.com/index.php?topic=149734.0'
$strFilePathAllKillerNoFillerRunNGunBatch = Join-Path $strSubfolderPath '_runNgun.txt'

if ((Test-Path $strFilePathAllKillerNoFillerRunNGunBatch) -ne $true) {
    Write-Error ('The All Killer No Filler Run ''n'' Gun batch file is missing. Please download it from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strURLAllKillerNoFillerRunNGun + "`n`n" + 'File Location:' + "`n" + $strFilePathAllKillerNoFillerRunNGunBatch)
    $boolErrorOccurred = $true
}

# "All Killer, No Filler" Beat 'em Up / Hack 'n' Slash Game List
# First forum post has an attachment in the form of a batch file
$strURLAllKillerNoFillerBeatEmUpHackNSlash = 'http://forum.arcadecontrols.com/index.php?topic=149833.0'
$strFilePathAllKillerNoFillerBeatEmUpHackNSlashBatch = Join-Path $strSubfolderPath '_beatNhack.txt'

if ((Test-Path $strFilePathAllKillerNoFillerBeatEmUpHackNSlashBatch) -ne $true) {
    Write-Error ('The All Killer No Filler Beat ''em Up / Hack ''n'' Slash batch file is missing. Please download it from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strURLAllKillerNoFillerBeatEmUpHackNSlash + "`n`n" + 'File Location:' + "`n" + $strFilePathAllKillerNoFillerBeatEmUpHackNSlashBatch)
    $boolErrorOccurred = $true
}

# "All Killer, No Filler" Platform Game List
# First forum post has an attachment in the form of a batch file
$strURLAllKillerNoFillerPlatform = 'http://forum.arcadecontrols.com/index.php?topic=150036.0'
$strFilePathAllKillerNoFillerPlatformBatch = Join-Path $strSubfolderPath '_platform.txt'

if ((Test-Path $strFilePathAllKillerNoFillerPlatformBatch) -ne $true) {
    Write-Error ('The All Killer No Filler Platform batch file is missing. Please download it from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strURLAllKillerNoFillerPlatform + "`n`n" + 'File Location:' + "`n" + $strFilePathAllKillerNoFillerPlatformBatch)
    $boolErrorOccurred = $true
}

# "All Killer, No Filler" Classics Game Lists
# First forum post has attachments in the form of batch files
$strURLAllKillerNoFillerClassics = 'http://forum.arcadecontrols.com/index.php?topic=150071.0'
$strFilePathAllKillerNoFillerClassicsHorizontalBatch = Join-Path $strSubfolderPath '_classicshor.txt'
$strFilePathAllKillerNoFillerClassicsVerticalBatch = Join-Path $strSubfolderPath '_classicsvert.txt'

if (((Test-Path $strFilePathAllKillerNoFillerClassicsHorizontalBatch) -ne $true) -or ((Test-Path $strFilePathAllKillerNoFillerClassicsVerticalBatch) -ne $true)) {
    Write-Error ('The All Killer No Filler Classics batch file(s) are missing. Please download them from the following URL and place them in the following locations.' + "`n`n" + 'URL: ' + $strURLAllKillerNoFillerClassics + "`n`n" + 'File Locations:' + "`n" + $strFilePathAllKillerNoFillerClassicsHorizontalBatch + "`n" + $strFilePathAllKillerNoFillerClassicsVerticalBatch)
    $boolErrorOccurred = $true
}

if ($boolErrorOccurred -eq $false) {
    # We have all the files, let's do stuff

    $csvCurrentRomList = @()

    $strCurrentFilePath = $strFilePathAllKillerNoFillerSHMUPSHorizontalBatch
    $strCurrentFileCategory = 'Shoot-Em-Up'
    $strCurrentFileScreenOrientation = 'Horizontal'
    Merge-AllKillerNoFillerFile ([ref]$csvCurrentRomList) $strCurrentFilePath $strCurrentFileCategory $strCurrentFileScreenOrientation

    $strCurrentFilePath = $strFilePathAllKillerNoFillerSHMUPSVerticalBatch
    $strCurrentFileCategory = 'Shoot-Em-Up'
    $strCurrentFileScreenOrientation = 'Vertical'
    Merge-AllKillerNoFillerFile ([ref]$csvCurrentRomList) $strCurrentFilePath $strCurrentFileCategory $strCurrentFileScreenOrientation

    $strCurrentFilePath = $strFilePathAllKillerNoFillerVSFightersBatch
    $strCurrentFileCategory = 'VS Fighter'
    $strCurrentFileScreenOrientation = 'Horizontal'
    Merge-AllKillerNoFillerFile ([ref]$csvCurrentRomList) $strCurrentFilePath $strCurrentFileCategory $strCurrentFileScreenOrientation

    $strCurrentFilePath = $strFilePathAllKillerNoFillerSportsBatch
    $strCurrentFileCategory = 'Sports'
    $strCurrentFileScreenOrientation = 'Horizontal'
    Merge-AllKillerNoFillerFile ([ref]$csvCurrentRomList) $strCurrentFilePath $strCurrentFileCategory $strCurrentFileScreenOrientation

    $strCurrentFilePath = $strFilePathAllKillerNoFillerPuzzleBatch
    $strCurrentFileCategory = 'Puzzle'
    $strCurrentFileScreenOrientation = 'Horizontal'
    Merge-AllKillerNoFillerFile ([ref]$csvCurrentRomList) $strCurrentFilePath $strCurrentFileCategory $strCurrentFileScreenOrientation

    $strCurrentFilePath = $strFilePathAllKillerNoFillerRunNGunBatch
    $strCurrentFileCategory = 'Run ''n'' Gun'
    $strCurrentFileScreenOrientation = 'Horizontal'
    Merge-AllKillerNoFillerFile ([ref]$csvCurrentRomList) $strCurrentFilePath $strCurrentFileCategory $strCurrentFileScreenOrientation

    $strCurrentFilePath = $strFilePathAllKillerNoFillerBeatEmUpHackNSlashBatch
    $strCurrentFileCategory = 'Beat ''em Up / Hack ''n'' Slash'
    $strCurrentFileScreenOrientation = 'Horizontal'
    Merge-AllKillerNoFillerFile ([ref]$csvCurrentRomList) $strCurrentFilePath $strCurrentFileCategory $strCurrentFileScreenOrientation

    $strCurrentFilePath = $strFilePathAllKillerNoFillerPlatformBatch
    $strCurrentFileCategory = 'Platformer'
    $strCurrentFileScreenOrientation = 'Horizontal'
    Merge-AllKillerNoFillerFile ([ref]$csvCurrentRomList) $strCurrentFilePath $strCurrentFileCategory $strCurrentFileScreenOrientation

    $strCurrentFilePath = $strFilePathAllKillerNoFillerClassicsHorizontalBatch
    $strCurrentFileCategory = 'Classic'
    $strCurrentFileScreenOrientation = 'Horizontal'
    Merge-AllKillerNoFillerFile ([ref]$csvCurrentRomList) $strCurrentFilePath $strCurrentFileCategory $strCurrentFileScreenOrientation

    $strCurrentFilePath = $strFilePathAllKillerNoFillerClassicsVerticalBatch
    $strCurrentFileCategory = 'Classic'
    $strCurrentFileScreenOrientation = 'Vertical'
    Merge-AllKillerNoFillerFile ([ref]$csvCurrentRomList) $strCurrentFilePath $strCurrentFileCategory $strCurrentFileScreenOrientation

    if ($strTargetSystem -eq 'Raspberry Pi') {
        # Add Star Gladiator (starglad) because the Raspberry Pi may not be fast enough to run its (superior) sequel
        $strThisROMName = 'starglad'
        $strCurrentFileCategory = 'VS Fighter'
        $strCurrentFileScreenOrientation = 'Horizontal'
        Merge-ROMManuallyOntoAllKillerNoFillerList ([ref]$csvCurrentRomList) $strThisROMName $strCurrentFileCategory $strCurrentFileScreenOrientation

        # Add Tekken 2 (tekken2) because the Raspberry Pi may not be fast enough to run its (superior) sequel
        $strThisROMName = 'tekken2'
        $strCurrentFileCategory = 'VS Fighter'
        $strCurrentFileScreenOrientation = 'Horizontal'
        Merge-ROMManuallyOntoAllKillerNoFillerList ([ref]$csvCurrentRomList) $strThisROMName $strCurrentFileCategory $strCurrentFileScreenOrientation
    }

    $csvCurrentRomList | Sort-Object -Property @('ROM') | Export-Csv $strCSVOutputFile -NoTypeInformation
}
