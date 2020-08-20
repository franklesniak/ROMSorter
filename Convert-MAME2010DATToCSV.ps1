# Convert-MAME2010DATToCSV.ps1
# Downloads the MAME 2010 DAT in XML format from Github, analyzes it, and stores the extracted
# data and associated insights in a CSV.

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
$strURL = 'https://github.com/libretro/mame2010-libretro/raw/master/metadata/mame2010.xml'

$strSubfolderPath = Join-Path '.' 'MAME_2010_Resources'
$strLocalXMLFilePath = $null

# Uncomment the following line if you prefer that the script use a local copy of the
#    MAME 2010 DAT file instead of having to download it from GitHub:
# $strLocalXMLFilePath = Join-Path $strSubfolderPath 'mame2010.xml'

$strOutputFilePath = Join-Path '.' 'MAME_2010_DAT.csv'
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

# Get the MAME 2010 DAT
if ($null -eq $strLocalXMLFilePath) {
    $strContent = Invoke-WebRequest -Uri $strURL
} else {
    if ((Test-Path $strLocalXMLFilePath) -ne $true) {
        Write-Error ('The MAME 2010 DAT file is missing. Please download it from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strURL + "`n`n" + 'File Location:' + "`n" + $strLocalXMLFilePath)
        break
    }
    $strAbsoluteXMLFilePath = (Resolve-Path $strLocalXMLFilePath).Path
    $strContent = [System.IO.File]::ReadAllText($strAbsoluteXMLFilePath)
}

# Convert it to XML
$xmlMAME2010 = [xml]$strContent

# Create a hashtable of game information for rapid lookup by name
$hashtableMAME2010 = New-BackwardCompatibleCaseInsensitiveHashtable
@($xmlMAME2010.mame.game) | ForEach-Object {
    $game = $_
    $hashtableMAME2010.Add($game.name, $game)
}

# Create an array of the types of controls
$arrInputTypes = @()
@($xmlMAME2010.mame.game) | ForEach-Object {
    $game = $_
    if ($null -ne $game.input) {
        @($game.input) | ForEach-Object {
            $inputFromXML = $_
            if ($null -ne $inputFromXML.control) {
                @($inputFromXML.control) | ForEach-Object {
                    $control = $_
                    if ($arrInputTypes -notcontains $control.type) {
                        $arrInputTypes += $control.type
                    }
                }
            }
        }
    }
}

# Translate legacy control types to updates ones used by newer versions of MAME
$arrControlsTotal = $arrInputTypes | ForEach-Object {
    $strInputType = $_
    switch ($strInputType) {
        'doublejoy2way' { $strAdjustedInputType = 'doublejoy_2wayhorizontal_2wayhorizontal' }
        'vdoublejoy2way' { $strAdjustedInputType = 'doublejoy_2wayvertical_2wayvertical' }
        'doublejoy4way' { $strAdjustedInputType = 'doublejoy_4way_4way' }
        'doublejoy8way' { $strAdjustedInputType = 'doublejoy_8way_8way' }
        'joy2way' { $strAdjustedInputType = 'joy_2wayhorizontal' }
        'vjoy2way' { $strAdjustedInputType = 'joy_2wayvertical' }
        'joy4way' { $strAdjustedInputType = 'joy_4way' }
        'joy8way' { $strAdjustedInputType = 'joy_8way' }
        default { $strAdjustedInputType = $strInputType }
    }
    $strAdjustedInputType
} | Select-Object -Unique | Sort-Object

# Create a hashtable used to associate the number of each type of input required for player 1
$hashtableInputCountsForPlayerOne = New-BackwardCompatibleCaseInsensitiveHashtable
$arrControlsTotal | ForEach-Object {
    $strInputType = $_
    $hashtableInputCountsForPlayerOne.Add($strInputType, 0)
}

$arrCSVMAME2010 = @($xmlMAME2010.mame.game) | ForEach-Object {
    $game = $_

    # Reset control counts
    $arrControlsTotal | ForEach-Object {
        $strInputType = $_
        $hashtableInputCountsForPlayerOne.Item($strInputType) = 0
    }

    $PSCustomObject = New-Object PSCustomObject
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'ROMName' -Value $game.name
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_ROMName' -Value $game.name
    if ($null -eq $game.description) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_ROMDisplayName' -Value ''
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_ROMDisplayName' -Value $game.description
    }
    if ($null -eq $game.manufacturer) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_Manufacturer' -Value ''
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_Manufacturer' -Value $game.manufacturer
    }
    if ($null -eq $game.year) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_Year' -Value ''
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_Year' -Value $game.year
    }
    if ($null -eq $game.cloneof) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_CloneOf' -Value ''
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_CloneOf' -Value $game.cloneof
    }
    if (($null -eq $game.isbios) -or ($game.isbios -eq 'no')) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_IsBIOSROM' -Value 'False'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_IsBIOSROM' -Value 'True'
    }

    $boolROMPackageContainsROMs = $false
    $boolROMPackageContainsCHD = $false
    $boolROMFunctional = Test-MachineCompletelyFunctionalRecursively ([ref]$boolROMPackageContainsROMs) ([ref]$boolROMPackageContainsCHD) ($game.name) ([ref]$hashtableMAME2010)

    if ($boolROMFunctional -eq $true) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_FunctionalROMPackage' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_FunctionalROMPackage' -Value 'False'
    }

    if ($boolROMPackageContainsROMs -eq $true) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_ROMFilesPartOfPackage' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_ROMFilesPartOfPackage' -Value 'False'
    }

    if ($boolROMPackageContainsCHD -eq $true) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_CHDsPartOfPackage' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_CHDsPartOfPackage' -Value 'False'
    }

    $boolSamplePresent = $false
    if ($null -ne $game.sample) {
        @($game.sample) | ForEach-Object {
            $boolSamplePresent = $true
        }
    }

    if ($boolSamplePresent -eq $true) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_SoundSamplesPartOfPackage' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_SoundSamplesPartOfPackage' -Value 'False'
    }

    if ($null -eq $game.display) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_DisplayCount' -Value '0'
        $intPrimaryDisplayIndex = -1
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_DisplayCount' -Value ([string](@($game.display).Count))
        $intPrimaryDisplayIndex = (@($game.display).Count) - 1
    }

    if ($intPrimaryDisplayIndex -gt 0) {
        # Multiple displays were present; find the primary one
        $intPrimaryDisplayIndex = 0
        $intMaxResolution = 0

        for ($intCounterA = 0; $intCounterA -lt @($game.display).Count; $intCounterA++) {
            $intCurrentDisplayWidth = [int](@($game.display)[$intCounterA].width)
            $intCurrentDisplayHeight = [int](@($game.display)[$intCounterA].height)
            $intCurrentResolution = $intCurrentDisplayWidth * $intCurrentDisplayHeight
            if ($intCurrentResolution -gt $intMaxResolution) {
                $intMaxResolution = $intCurrentResolution
                $intPrimaryDisplayIndex = $intCounterA
            }
        }
    }

    if ($intPrimaryDisplayIndex -ge 0) {
        if ((@($game.display)[$intPrimaryDisplayIndex].rotate -eq '90') -or (@($game.display)[$intPrimaryDisplayIndex].rotate -eq '270')) {
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_PrimaryDisplayOrientation' -Value 'Vertical'
            $intCurrentDisplayHeight = [int](@($game.display)[$intPrimaryDisplayIndex].width)
            $intCurrentDisplayWidth = [int](@($game.display)[$intPrimaryDisplayIndex].height)
        } else {
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_PrimaryDisplayOrientation' -Value 'Horizontal'
            $intCurrentDisplayWidth = [int](@($game.display)[$intPrimaryDisplayIndex].width)
            $intCurrentDisplayHeight = [int](@($game.display)[$intPrimaryDisplayIndex].height)
        }
        $doubleRefreshRate = [double](@($game.display)[$intPrimaryDisplayIndex].refresh)
        $strResolution = ([string]$intCurrentDisplayWidth) + 'x' + ([string]$intCurrentDisplayHeight) + '@' + ([string]$doubleRefreshRate) + 'Hz'
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_PrimaryDisplayResolution' -Value $strResolution
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_PrimaryDisplayOrientation' -Value 'N/A'
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_PrimaryDisplayResolution' -Value 'N/A'
    }

    if ($null -ne $game.sound) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_ROMPackageHasSound' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_ROMPackageHasSound' -Value 'False'
    }

    $strNumPlayers = 'N/A'
    $strNumButtons = 'N/A'
    if ($null -ne $game.input) {
        @($game.input) | ForEach-Object {
            $inputFromXML = $_
            if ($null -ne $inputFromXML.players) {
                if ($strNumPlayers -eq 'N/A') {
                    $strNumPlayers = '0'
                }
                if (([int]($inputFromXML.players)) -gt ([int]$strNumPlayers)) {
                    $strNumPlayers = $inputFromXML.players
                }
            }
            if ($null -ne $inputFromXML.buttons) {
                if ($strNumButtons -eq 'N/A') {
                    $strNumButtons = '0'
                }
                if (([int]($inputFromXML.buttons)) -gt ([int]$strNumButtons)) {
                    $strNumButtons = $inputFromXML.buttons
                }
            }
            if ($null -ne $inputFromXML.control) {
                @($inputFromXML.control) | ForEach-Object {
                    $control = $_
                    if ($null -ne $control.type) {
                        $strInputType = $control.type
                        switch ($strInputType) {
                            'doublejoy2way' { $strAdjustedInputType = 'doublejoy_2wayhorizontal_2wayhorizontal' }
                            'vdoublejoy2way' { $strAdjustedInputType = 'doublejoy_2wayvertical_2wayvertical' }
                            'doublejoy4way' { $strAdjustedInputType = 'doublejoy_4way_4way' }
                            'doublejoy8way' { $strAdjustedInputType = 'doublejoy_8way_8way' }
                            'joy2way' { $strAdjustedInputType = 'joy_2wayhorizontal' }
                            'vjoy2way' { $strAdjustedInputType = 'joy_2wayvertical' }
                            'joy4way' { $strAdjustedInputType = 'joy_4way' }
                            'joy8way' { $strAdjustedInputType = 'joy_8way' }
                            default { $strAdjustedInputType = $strInputType }
                        }
                        $hashtableInputCountsForPlayerOne.Item($strAdjustedInputType)++
                    }
                }
            }
        }
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_ROMPackageHasInput' -Value 'True'
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_NumberOfPlayers' -Value $strNumPlayers
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_NumberOfButtons' -Value $strNumButtons
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = $hashtableInputCountsForPlayerOne.Item($strInputType)
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME2010_P1_NumInputControls_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_ROMPackageHasInput' -Value 'False'
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_NumberOfPlayers' -Value $strNumPlayers
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_NumberOfButtons' -Value $strNumButtons
        $arrControlsTotal | ForEach-Object {
            $strInputType = $_
            $intNumControlsOfThisType = 0
            $PSCustomObject | Add-Member -MemberType NoteProperty -Name ('MAME2010_P1_NumInputControls_' + $strInputType) -Value ([string]$intNumControlsOfThisType)
        }
    }

    $boolFreePlaySupported = $false
    $arrSupportedCabinetTypes = @()
    if ($null -ne $game.dipswitch) {
        @($game.dipswitch) | ForEach-Object {
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
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_FreePlaySupported' -Value 'True'
    } else {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_FreePlaySupported' -Value 'False'
    }
    if ($arrSupportedCabinetTypes.Count -eq 0) {
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_CabinetTypes' -Value 'Unknown'
    } else {
        $strCabinetTypes = ($arrSupportedCabinetTypes | Sort-Object) -join ';'
        $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_CabinetTypes' -Value $strCabinetTypes
    }

    $strOverallStatus = 'Unknown'
    $strEmulationStatus = 'Unknown'
    $strColorStatus = 'Unknown'
    $strSoundStatus = 'Unknown'
    $strGraphicStatus = 'Unknown'
    $strCocktailStatus = 'Unknown'
    $strProtectionStatus = 'Unknown'
    $strSaveStateSupported = 'Unknown'
    $strPaletteSize = 'Unknown'
    if ($null -ne $game.driver) {
        @($game.driver) | ForEach-Object {
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

            switch ($driver.color) {
                'good' { $strTemp = 'Good' }
                'imperfect' { $strTemp = 'Imperfect' }
                'preliminary' { $strTemp = 'Preliminary' }
                default { $strTemp = $driver.color }
            }
            $strColorStatus = $strTemp

            switch ($driver.sound) {
                'good' { $strTemp = 'Good' }
                'imperfect' { $strTemp = 'Imperfect' }
                'preliminary' { $strTemp = 'Preliminary' }
                default { $strTemp = $driver.sound }
            }
            $strSoundStatus = $strTemp

            switch ($driver.graphic) {
                'good' { $strTemp = 'Good' }
                'imperfect' { $strTemp = 'Imperfect' }
                'preliminary' { $strTemp = 'Preliminary' }
                default { $strTemp = $driver.graphic }
            }
            $strGraphicStatus = $strTemp

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

            if ($null -ne $driver.protection) {
                switch ($driver.protection) {
                    'good' { $strTemp = 'Good' }
                    'imperfect' { $strTemp = 'Imperfect' }
                    'preliminary' { $strTemp = 'Preliminary' }
                    default { $strTemp = $driver.protection }
                }
                $strProtectionStatus = $strTemp
            } else {
                $strProtectionStatus = 'Not Specified'
            }

            if ($driver.savestate -eq 'supported') {
                $strSaveStateSupported = 'True'
            } else {
                $strSaveStateSupported = 'False'
            }

            $strPaletteSize = $driver.palettesize
        }
    }
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_OverallStatus' -Value $strOverallStatus
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_EmulationStatus' -Value $strEmulationStatus
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_ColorStatus' -Value $strColorStatus
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_SoundStatus' -Value $strSoundStatus
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_GraphicStatus' -Value $strGraphicStatus
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_CocktailStatus' -Value $strCocktailStatus
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_ProtectionStatus' -Value $strProtectionStatus
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_SaveStateSupported' -Value $strSaveStateSupported
    $PSCustomObject | Add-Member -MemberType NoteProperty -Name 'MAME2010_PaletteSize' -Value $strPaletteSize

    $PSCustomObject
}

$arrCSVMAME2010 | Sort-Object -Property @('ROMName') |
    Export-Csv -Path $strOutputFilePath -NoTypeInformation
