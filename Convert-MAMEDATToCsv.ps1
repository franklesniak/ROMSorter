# Convert-MAMEDATToCSV.ps1
# Analyzes the current version of MAME's DAT in XML format and stores the extracted data and
# associated insights in a CSV.

$strThisScriptVersionNumber = [version]'1.0.20211227.0'

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
$strLocalXMLFilePath = $null

#region Inputs
###############################################################################################
$strDownloadPageURL = 'https://www.mamedev.org/release.html'
$strURL = 'https://github.com/mamedev/mame/releases/download/mame0238/mame0238lx.zip'

$strSubfolderPath = Join-Path '.' 'MAME_Resources'

# Uncomment and configure the following line if you prefer that the script use a local copy of
#   the MAME DAT file instead of having to download it from GitHub:

# $strLocalXMLFilePath = Join-Path $strSubfolderPath 'mame.xml'

$strOutputFilePath = Join-Path '.' 'MAME_DAT.csv'

# Comment-out the following line if you prefer that the script operate silently.
$actionPreferenceNewVerbose = [System.Management.Automation.ActionPreference]::Continue
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

function Test-MachineCompletelyFunctionalRecursively {
    # This functions supports recursive ROM lookups in a MAME DAT to determine if a non-merged
    # romset containing this machine (ROM package) would be considered non-functional (i.e.,
    # having a baddump or nodump ROM or CHD, or runnable equal to 'no'). If the machine (ROM
    # package) in a non-merged romset is non-functional, this function returns $false;
    # otherwise, it returns $true
    #
    # The function takes four positional arguments.
    #
    # The first argument is a reference to a boolean variable. Before calling the function,
    # the boolean variable must be initialized to $false. After completion of the function, the
    # boolean variable is set to $true if this machine (ROM package), in a non-merged romset,
    # would contain at least one ROM file.
    #
    # The second argument is also a reference to a boolean variable, and before calling the
    # function, this boolean variable must also be initialized to $false. After completion of
    # the function, the boolean variable is set to $true if this machine (ROM package), in a
    # non-merged romset, would contain at least one CHD file.
    #
    # The third argument is a string containing the short name of the machine (ROM package).
    #
    # The fourth argument is a reference to a hashtable of all the ROM information obtained
    # from the DAT, indexed by the ROM name.
    #
    # Example:
    # $strROMName = 'mario'
    # $boolROMPackageContainsROMs = $false
    # $boolROMPackageContainsCHD = $false
    # $boolROMFunctional = Test-MachineCompletelyFunctionalRecursively ([ref]$boolROMPackageContainsROMs) ([ref]$boolROMPackageContainsCHD) $strROMName ([ref]$hashtableEmulatorDAT)

    $refBoolROMPackagePresent = $args[0]
    $refBoolCHDPresent = $args[1]
    $strThisROMName = $args[2]
    $refHashtableDAT = $args[3]

    $strThisFunctionVersionNumber = [version]'1.0.20200820.0'

    $game = ($refHashtableDAT.Value).Item($strThisROMName)
    $boolParentROMCompletelyFunctional = $true
    if ($null -ne $game.romof) {
        # This game has a parent ROM
        $boolParentROMCompletelyFunctional = Test-MachineCompletelyFunctionalRecursively $refBoolROMPackagePresent $refBoolCHDPresent ($game.romof) $refHashtableDAT
    }

    if ($boolParentROMCompletelyFunctional -eq $false) {
        $false
    } else {
        $boolCompletelyFunctionalROMPackage = $true

        if ($game.runnable -eq 'no') {
            $boolCompletelyFunctionalROMPackage = $false
        }

        if ($null -ne $game.rom) {
            @($game.rom) | ForEach-Object {
                $file = $_
                ($refBoolROMPackagePresent.Value) = $true
                $boolOptionalFile = $false
                if ($file.optional -eq 'yes') {
                    $boolOptionalFile = $true
                }
                if ($boolOptionalFile -eq $false) {
                    if (($file.status -eq 'baddump') -or ($file.status -eq 'nodump')) {
                        $boolCompletelyFunctionalROMPackage = $false
                    }
                }
            }
        }
        if ($null -ne $game.disk) {
            @($game.disk) | ForEach-Object {
                $file = $_
                ($refBoolCHDPresent.Value) = $true
                $boolOptionalFile = $false
                if ($file.optional -eq 'yes') {
                    $boolOptionalFile = $true
                }
                if ($boolOptionalFile -eq $false) {
                    if (($file.status -eq 'baddump') -or ($file.status -eq 'nodump')) {
                        $boolCompletelyFunctionalROMPackage = $false
                    }
                }
            }
        }
        $boolCompletelyFunctionalROMPackage
    }
}

function Convert-MAMEControlWaysToClearerOutputFormat {
    $strControlWaysFromDAT = $args[0]
    switch ($strControlWaysFromDAT) {
        '1' { $strAdjustedInputType = '1way' }
        '2' { $strAdjustedInputType = '2wayhorizontal' }
        '3 (half4)' { $strAdjustedInputType = '3way' }
        '4' { $strAdjustedInputType = '4way' }
        '5 (half8)' { $strAdjustedInputType = '5way' }
        '8' { $strAdjustedInputType = '8way' }
        '16' { $strAdjustedInputType = '16way' }
        'strange2' { $strAdjustedInputType = '2waystrange' }
        'vertical2' { $strAdjustedInputType = '2wayvertical' }
        default { $strAdjustedInputType = $strControlWaysFromDAT }
    }
    $strAdjustedInputType
}

$VerbosePreference = $actionPreferenceNewVerbose

# Get the MAME DAT
$arrCommands = @(Get-Command Expand-Archive)
$boolZIPExtractAvailable = ($arrCommands.Count -ge 1)
$arrCommands = @(Get-Command Invoke-WebRequest)
$boolInvokeWebRequestAvailable = ($arrCommands.Count -ge 1)
if ($null -eq $strLocalXMLFilePath -and $boolZIPExtractAvailable -and $boolInvokeWebRequestAvailable) {
    $VerbosePreference = $actionPreferenceFormerVerbose
    $arrModules = @(Get-Module PowerHTML -ListAvailable)
    $VerbosePreference = $actionPreferenceNewVerbose
    if ($arrModules.Count -eq 0) {
        Write-Warning 'It is recommended that you install the PowerHTML module using "Install-Module PowerHTML" before continuing. Doing so will allow this script to obtain the URL for the most-current DAT file automatically. Without PowerHTML, this script is using a potentially-outdated URL. Break out of ths script now to install PowerHTML, then re-run the script'
        $strEffectiveURL = $strURL
    } else {
        Write-Verbose ('Parsing site ' + $strDownloadPageURL + ' to dynamically obtain DAT download URL...')
        $arrLoadedModules = @(Get-Module PowerHTML)
        if ($arrLoadedModules.Count -eq 0) {
            $VerbosePreference = $actionPreferenceFormerVerbose
            Import-Module PowerHTML
            $VerbosePreference = $actionPreferenceNewVerbose
        }
        $strNextDownloadPageURL = $strDownloadPageURL
        $HtmlNodeDownloadPage = ConvertFrom-Html -URI $strNextDownloadPageURL
        $arrNodes = @($HtmlNodeDownloadPage.SelectNodes('//a[@href]') | Where-Object { $_.InnerText.ToLower().Contains('lx.zip') })
        if ($arrNodes.Count -eq 0) {
            Write-Error ('Failed to download the MAME DAT file. Please download the file that looks like mame*lx.zip from the following URL, extract the ZIP, and place it in the following location.' + "`n`n" + 'URL: ' + $strDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalXMLFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalXMLFilePath to point to the path of the downloaded and extracted XML file.')
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
    Write-Verbose ('Downloading compressed DAT from ' + $strEffectiveURL + '...')
    $VerbosePreference = $actionPreferenceFormerVerbose
    Invoke-WebRequest -Uri $strEffectiveURL -OutFile (Join-Path $strSubfolderPath mamelx.zip)
    $VerbosePreference = $actionPreferenceNewVerbose

    if (Test-Path (Join-Path $strSubfolderPath mamelx.zip)) {
        # Successful download
        Write-Verbose 'Extracting DAT from compressed ZIP...'
        # TODO: Create backward compatible alternative to Expand-Archive
        Expand-Archive -Path (Join-Path $strSubfolderPath mamelx.zip) -DestinationPath $strSubfolderPath -Force
        $fileInfoExtractedXML = Get-ChildItem $strSubfolderPath | Where-Object { $_.Name.Length -ge 5 } | `
            Where-Object { $_.Name.Substring(($_.Name.Length - 4), 4).ToLower() -eq '.xml' } | `
            Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1
        $strAbsoluteXMLFilePath = $fileInfoExtractedXML.FullName
        Write-Verbose ('Loading DAT into memory and converting it to XML object...')
        $strContent = [System.IO.File]::ReadAllText($strAbsoluteXMLFilePath)
    } else {
        Write-Error ('Failed to download the MAME DAT file. Please download the file that looks like mame*lx.zip from the following URL, extract the ZIP, and place it in the following location.' + "`n`n" + 'URL: ' + $strDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalXMLFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalXMLFilePath to point to the path of the downloaded and extracted XML file.')
        break
    }
} else {
    if ((Test-Path $strLocalXMLFilePath) -ne $true) {
        Write-Error ('The MAME DAT file is missing. Please download the file that looks like mame*lx.zip from the following URL, extract the ZIP, and place it in the following location.' + "`n`n" + 'URL: ' + $strDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalXMLFilePath)
        break
    }
    $strAbsoluteXMLFilePath = (Resolve-Path $strLocalXMLFilePath).Path
    Write-Verbose ('Loading DAT into memory and converting it to XML object...')
    $strContent = [System.IO.File]::ReadAllText($strAbsoluteXMLFilePath)
}

# Convert it to XML
$xmlMAME = [xml]$strContent

Write-Verbose ('Creating a hashtable of ROM package information for rapid lookup by name...')
$hashtableMAME = New-BackwardCompatibleCaseInsensitiveHashtable
@($xmlMAME.mame.machine) | ForEach-Object {
    $machine = $_
    $hashtableMAME.Add($machine.name, $machine)
}

Write-Verbose ('Creating a array to act as a dictionary of the different types of controls available in this DAT...')
$arrInputTypes = @()
@($xmlMAME.mame.machine) | ForEach-Object {
    $machine = $_
    if ($null -ne $machine.input) {
        @($machine.input) | ForEach-Object {
            $inputFromXML = $_
            if ($null -ne $inputFromXML.control) {
                @($inputFromXML.control) | ForEach-Object {
                    $control = $_
                    $strControlString = $control.type
                    if ($null -ne $control.ways) {
                        if ($control.ways -ne '') {
                            $strControlString = $strControlString + '_' + (Convert-MAMEControlWaysToClearerOutputFormat $control.ways)
                            if ($null -ne $control.ways2) {
                                if ($control.ways2 -ne '') {
                                    $strControlString = $strControlString + '_' + (Convert-MAMEControlWaysToClearerOutputFormat $control.ways2)
                                    if ($null -ne $control.ways3) {
                                        if ($control.ways3 -ne '') {
                                            $strControlString = $strControlString + '_' + (Convert-MAMEControlWaysToClearerOutputFormat $control.ways3)
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if ($arrInputTypes -notcontains $strControlString) {
                        $arrInputTypes += $strControlString
                    }
                }
            }
        }
    }
}

$arrControlsTotal = $arrInputTypes | Sort-Object

# Create a hashtable used to associate the number of each type of input required across all players
$hashtableInputCountsForAllPlayers = New-BackwardCompatibleCaseInsensitiveHashtable
$arrControlsTotal | ForEach-Object {
    $strInputType = $_
    $hashtableInputCountsForAllPlayers.Add($strInputType, 0)
}

# Create a hashtable used to associate the number of each type of input required for player 1
$hashtableInputCountsForPlayerOne = New-BackwardCompatibleCaseInsensitiveHashtable
$arrControlsTotal | ForEach-Object {
    $strInputType = $_
    $hashtableInputCountsForPlayerOne.Add($strInputType, 0)
}

# Create a hashtable used to associate the number of each type of input required for player 2
$hashtableInputCountsForPlayerTwo = New-BackwardCompatibleCaseInsensitiveHashtable
$arrControlsTotal | ForEach-Object {
    $strInputType = $_
    $hashtableInputCountsForPlayerTwo.Add($strInputType, 0)
}

# Create a hashtable used to associate the number of each type of input required for player 3
$hashtableInputCountsForPlayerThree = New-BackwardCompatibleCaseInsensitiveHashtable
$arrControlsTotal | ForEach-Object {
    $strInputType = $_
    $hashtableInputCountsForPlayerThree.Add($strInputType, 0)
}

# Create a hashtable used to associate the number of each type of input required for player 4
$hashtableInputCountsForPlayerFour = New-BackwardCompatibleCaseInsensitiveHashtable
$arrControlsTotal | ForEach-Object {
    $strInputType = $_
    $hashtableInputCountsForPlayerFour.Add($strInputType, 0)
}

Write-Verbose ('Processing ROM packages...')

$intTotalROMPackages = @($xmlMAME.mame.machine).Count
$intCurrentROMPackage = 1
$timeDateStartOfProcessing = Get-Date

$arrCSVMAME = @($xmlMAME.mame.machine) | ForEach-Object {
    $machine = $_

    if ($intCurrentROMPackage -ge 101) {
        $timeDateCurrent = Get-Date
        $timeSpanElapsed = $timeDateCurrent - $timeDateStartOfProcessing
        $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / ($intCurrentROMPackage - 1) * $intTotalROMPackages
        $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
        $doublePercentComplete = ($intCurrentROMPackage - 1) / $intTotalROMPackages * 100
        Write-Progress -Activity 'Processing MAME ROM Packages' -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
    }

    # Reset control counts
    $arrControlsTotal | ForEach-Object {
        $strInputType = $_
        $hashtableInputCountsForAllPlayers.Item($strInputType) = 0
        $hashtableInputCountsForPlayerOne.Item($strInputType) = 0
        $hashtableInputCountsForPlayerTwo.Item($strInputType) = 0
        $hashtableInputCountsForPlayerThree.Item($strInputType) = 0
        $hashtableInputCountsForPlayerFour.Item($strInputType) = 0
    }

    $PSCustomObject = New-Object PSCustomObject
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'ROMName' -Value $machine.name
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_ROMName' -Value $machine.name
    if ($null -eq $machine.description) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_ROMDisplayName' -Value ''
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_ROMDisplayName' -Value $machine.description
    }
    if ($null -eq $machine.manufacturer) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_Manufacturer' -Value ''
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_Manufacturer' -Value $machine.manufacturer
    }
    if ($null -eq $machine.year) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_Year' -Value ''
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_Year' -Value $machine.year
    }
    if ($null -eq $machine.cloneof) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_CloneOf' -Value ''
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_CloneOf' -Value $machine.cloneof
    }
    if (($null -eq $machine.isbios) -or ($machine.isbios -eq 'no')) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_IsBIOSROM' -Value 'False'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_IsBIOSROM' -Value 'True'
    }
    if (($null -eq $machine.isdevice) -or ($machine.isdevice -eq 'no')) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_IsDeviceROM' -Value 'False'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_IsDeviceROM' -Value 'True'
    }
    if (($null -eq $machine.ismechanical) -or ($machine.ismechanical -eq 'no')) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_IsMechanicalROM' -Value 'False'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_IsMechanicalROM' -Value 'True'
    }

    $boolROMPackageContainsROMs = $false
    $boolROMPackageContainsCHD = $false
    $boolROMFunctional = Test-MachineCompletelyFunctionalRecursively ([ref]$boolROMPackageContainsROMs) ([ref]$boolROMPackageContainsCHD) ($machine.name) ([ref]$hashtableMAME)

    if ($boolROMFunctional -eq $true) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_FunctionalROMPackage' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_FunctionalROMPackage' -Value 'False'
    }

    if ($boolROMPackageContainsROMs -eq $true) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_ROMFilesPartOfPackage' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_ROMFilesPartOfPackage' -Value 'False'
    }

    if ($boolROMPackageContainsCHD -eq $true) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_CHDsPartOfPackage' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_CHDsPartOfPackage' -Value 'False'
    }

    $boolSamplePresent = $false
    if ($null -ne $machine.sample) {
        @($machine.sample) | ForEach-Object {
            $boolSamplePresent = $true
        }
    }

    if ($boolSamplePresent -eq $true) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_SoundSamplesPartOfPackage' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_SoundSamplesPartOfPackage' -Value 'False'
    }

    if ($null -eq $machine.display) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_DisplayCount' -Value '0'
        $intPrimaryDisplayIndex = -1
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_DisplayCount' -Value ([string](@($machine.display).Count))
        $intPrimaryDisplayIndex = (@($machine.display).Count) - 1
    }

    if ($intPrimaryDisplayIndex -gt 0) {
        # Multiple displays were present; find the primary one
        $intPrimaryDisplayIndex = 0
        $intMaxResolution = 0

        for ($intCounterA = 0; $intCounterA -lt @($machine.display).Count; $intCounterA++) {
            $intCurrentDisplayWidth = [int](@($machine.display)[$intCounterA].width)
            $intCurrentDisplayHeight = [int](@($machine.display)[$intCounterA].height)
            $intCurrentResolution = $intCurrentDisplayWidth * $intCurrentDisplayHeight
            if ($intCurrentResolution -gt $intMaxResolution) {
                $intMaxResolution = $intCurrentResolution
                $intPrimaryDisplayIndex = $intCounterA
            }
        }
    }

    if ($intPrimaryDisplayIndex -ge 0) {
        if ((@($machine.display)[$intPrimaryDisplayIndex].rotate -eq '90') -or (@($machine.display)[$intPrimaryDisplayIndex].rotate -eq '270')) {
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_PrimaryDisplayOrientation' -Value 'Vertical'
            $intCurrentDisplayHeight = [int](@($machine.display)[$intPrimaryDisplayIndex].width)
            $intCurrentDisplayWidth = [int](@($machine.display)[$intPrimaryDisplayIndex].height)
        } else {
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_PrimaryDisplayOrientation' -Value 'Horizontal'
            $intCurrentDisplayWidth = [int](@($machine.display)[$intPrimaryDisplayIndex].width)
            $intCurrentDisplayHeight = [int](@($machine.display)[$intPrimaryDisplayIndex].height)
        }
        $doubleRefreshRate = [double](@($machine.display)[$intPrimaryDisplayIndex].refresh)
        $strResolution = ([string]$intCurrentDisplayWidth) + 'x' + ([string]$intCurrentDisplayHeight) + '@' + ([string]$doubleRefreshRate) + 'Hz'
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_PrimaryDisplayResolution' -Value $strResolution
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_PrimaryDisplayOrientation' -Value 'N/A'
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_PrimaryDisplayResolution' -Value 'N/A'
    }

    if ($null -ne $machine.sound) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_ROMPackageHasSound' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_ROMPackageHasSound' -Value 'False'
    }

    $strNumPlayers = 'N/A'
    $strNumButtonsTotal = 'N/A'
    $strNumButtonsPlayerOne = 'N/A'
    $strNumButtonsPlayerTwo = 'N/A'
    $strNumButtonsPlayerThree = 'N/A'
    $strNumButtonsPlayerFour = 'N/A'
    if ($null -ne $machine.input) {
        @($machine.input) | ForEach-Object {
            $inputFromXML = $_
            if ($null -ne $inputFromXML.players) {
                if ($strNumPlayers -eq 'N/A') {
                    $strNumPlayers = '0'
                }
                if (([int]($inputFromXML.players)) -gt ([int]$strNumPlayers)) {
                    $strNumPlayers = $inputFromXML.players
                }
            }
            if ($null -ne $inputFromXML.control) {
                @($inputFromXML.control) | ForEach-Object {
                    $control = $_

                    if ($null -ne $control.player) {
                        $intCurrentPlayer = [int]$control.player
                    } else {
                        $intCurrentPlayer = 0
                    }

                    if ($null -ne $control.buttons) {
                        if ($strNumButtonsTotal -eq 'N/A') {
                            $strNumButtonsTotal = '0'
                        }
                        $strNumButtonsTotal = [string]([int]($strNumButtonsTotal) + [int]($control.buttons))
                    }

                    $strControlString = $control.type
                    if ($null -ne $control.ways) {
                        if ($control.ways -ne '') {
                            $strControlString = $strControlString + '_' + (Convert-MAMEControlWaysToClearerOutputFormat $control.ways)
                            if ($null -ne $control.ways2) {
                                if ($control.ways2 -ne '') {
                                    $strControlString = $strControlString + '_' + (Convert-MAMEControlWaysToClearerOutputFormat $control.ways2)
                                    if ($null -ne $control.ways3) {
                                        if ($control.ways3 -ne '') {
                                            $strControlString = $strControlString + '_' + (Convert-MAMEControlWaysToClearerOutputFormat $control.ways3)
                                        }
                                    }
                                }
                            }
                        }
                    }

                    $hashtableInputCountsForAllPlayers.Item($strControlString)++

                    if ($intCurrentPlayer -eq 1) {
                        if ($null -ne $control.buttons) {
                            if ($strNumButtonsPlayerOne -eq 'N/A') {
                                $strNumButtonsPlayerOne = '0'
                            }
                            $strNumButtonsPlayerOne = [string]([int]($strNumButtonsPlayerOne) + [int]($control.buttons))
                        }

                        $hashtableInputCountsForPlayerOne.Item($strControlString)++
                    } elseif ($intCurrentPlayer -eq 2) {
                        if ($null -ne $control.buttons) {
                            if ($strNumButtonsPlayerTwo -eq 'N/A') {
                                $strNumButtonsPlayerTwo = '0'
                            }
                            $strNumButtonsPlayerTwo = [string]([int]($strNumButtonsPlayerTwo) + [int]($control.buttons))
                        }

                        $hashtableInputCountsForPlayerTwo.Item($strControlString)++
                    } elseif ($intCurrentPlayer -eq 3) {
                        if ($null -ne $control.buttons) {
                            if ($strNumButtonsPlayerThree -eq 'N/A') {
                                $strNumButtonsPlayerThree = '0'
                            }
                            $strNumButtonsPlayerThree = [string]([int]($strNumButtonsPlayerThree) + [int]($control.buttons))
                        }

                        $hashtableInputCountsForPlayerThree.Item($strControlString)++
                    } elseif ($intCurrentPlayer -eq 1) {
                        if ($null -ne $control.buttons) {
                            if ($strNumButtonsPlayerFour -eq 'N/A') {
                                $strNumButtonsPlayerFour = '0'
                            }
                            $strNumButtonsPlayerFour = [string]([int]($strNumButtonsPlayerFour) + [int]($control.buttons))
                        }

                        $hashtableInputCountsForPlayerFour.Item($strControlString)++
                    }
                }
            }
        }
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_ROMPackageHasInput' -Value 'True'
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_NumberOfPlayers' -Value $strNumPlayers
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_NumberOfButtonsAcrossAllPlayers' -Value $strNumButtonsTotal
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = $hashtableInputCountsForAllPlayers.Item($strInputType)
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME_NumInputControlsAcrossAllPlayers_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_P1_NumberOfButtons' -Value $strNumButtonsPlayerOne
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = $hashtableInputCountsForPlayerOne.Item($strInputType)
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME_P1_NumInputControls_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_P2_NumberOfButtons' -Value $strNumButtonsPlayerTwo
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = $hashtableInputCountsForPlayerTwo.Item($strInputType)
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME_P2_NumInputControls_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_P3_NumberOfButtons' -Value $strNumButtonsPlayerThree
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = $hashtableInputCountsForPlayerThree.Item($strInputType)
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME_P3_NumInputControls_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_P4_NumberOfButtons' -Value $strNumButtonsPlayerFour
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = $hashtableInputCountsForPlayerFour.Item($strInputType)
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME_P4_NumInputControls_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_ROMPackageHasInput' -Value 'False'
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_NumberOfPlayers' -Value $strNumPlayers
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_NumberOfButtonsAcrossAllPlayers' -Value $strNumButtonsTotal
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = 0
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME_NumInputControlsAcrossAllPlayers_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_P1_NumberOfButtons' -Value $strNumButtonsPlayerOne
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = 0
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME_P1_NumInputControls_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_P2_NumberOfButtons' -Value $strNumButtonsPlayerTwo
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = 0
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME_P2_NumInputControls_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_P3_NumberOfButtons' -Value $strNumButtonsPlayerThree
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = 0
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME_P3_NumInputControls_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_P4_NumberOfButtons' -Value $strNumButtonsPlayerFour
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = 0
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME_P4_NumInputControls_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
    }

    $boolFreePlaySupported = $false
    $arrSupportedCabinetTypes = @()
    if ($null -ne $machine.dipswitch) {
        @($machine.dipswitch) | ForEach-Object {
            $dipswitch = $_
            if ($dipswitch.name -eq 'Free Play') {
                $boolFreePlaySupported = $true
            }
            if ($dipswitch.name -eq 'Cabinet') {
                if ($null -ne $dipswitch.dipvalue) {
                    @($dipswitch.dipvalue) | ForEach-Object {
                        $dipvalue = $_
                        if ($arrSupportedCabinetTypes -notcontains $dipvalue.name) {
                            $arrSupportedCabinetTypes += $dipvalue.name
                        }
                    }
                }
            }
        }
    }
    if ($boolFreePlaySupported) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_FreePlaySupported' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_FreePlaySupported' -Value 'False'
    }
    if ($arrSupportedCabinetTypes.Count -eq 0) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_CabinetTypes' -Value 'Unknown'
    } else {
        $strCabinetTypes = ($arrSupportedCabinetTypes | Sort-Object) -join ';'
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_CabinetTypes' -Value $strCabinetTypes
    }

    $strOverallStatus = 'Unknown'
    $strEmulationStatus = 'Unknown'
    $strCocktailStatus = 'Unknown'
    $strSaveStateSupported = 'Unknown'
    if ($null -ne $machine.driver) {
        @($machine.driver) | ForEach-Object {
            $driver = $_

            switch ($driver.status) {
                'good' { $strTemp = 'Good' }
                'imperfect' { $strTemp = 'Imperfect' }
                'preliminary' { $strTemp = 'Preliminary' }
                default { $strTemp = $driver.status }
            }
            $strOverallStatus = $strTemp

            switch ($driver.emulation) {
                'good' { $strTemp = 'Good' }
                'imperfect' { $strTemp = 'Imperfect' }
                'preliminary' { $strTemp = 'Preliminary' }
                default { $strTemp = $driver.status }
            }
            $strEmulationStatus = $strTemp

            if ($null -ne $driver.cocktail) {
                switch ($driver.cocktail) {
                    'good' { $strTemp = 'Good' }
                    'imperfect' { $strTemp = 'Imperfect' }
                    'preliminary' { $strTemp = 'Preliminary' }
                    default { $strTemp = $driver.cocktail }
                }
                $strCocktailStatus = $strTemp
            } else {
                $strCocktailStatus = 'Not Specified'
            }

            if ($driver.savestate -eq 'supported') {
                $strSaveStateSupported = 'True'
            } else {
                $strSaveStateSupported = 'False'
            }

            $strPaletteSize = $driver.palettesize
        }
    }
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_OverallStatus' -Value $strOverallStatus
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_EmulationStatus' -Value $strEmulationStatus
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_CocktailStatus' -Value $strCocktailStatus
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME_SaveStateSupported' -Value $strSaveStateSupported

    $PSCustomObject
    
    $intCurrentROMPackage++
}

Write-Verbose ('Exporting results to CSV: ' + $strOutputFilePath)
$arrCSVMAME | Sort-Object -Property @('ROMName') |
    Export-Csv -Path $strOutputFilePath -NoTypeInformation
$VerbosePreference = $actionPreferenceFormerVerbose
