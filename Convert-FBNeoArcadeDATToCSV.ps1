# Convert-FBNeoArcadeDATToCSV.ps1
# Analyzes the current version of Final Burn Neo (FBNeo)'s arcade-only DAT in XML format and
# stores the extracted data and associated insights in a CSV.

$strThisScriptVersionNumber = [version]'1.2.20211230.0'

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
$strDownloadPageURL = 'https://github.com/libretro/FBNeo'
$strURL = 'https://raw.githubusercontent.com/libretro/FBNeo/master/dats/FinalBurn%20Neo%20(ClrMame%20Pro%20XML%2C%20Arcade%20only).dat'

$strSubfolderPath = Join-Path '.' 'FBNeo_Resources'

# Uncomment and configure the following line if you prefer that the script use a local copy of
#   the FBNeo DAT file instead of having to download it from GitHub:

# $strLocalXMLFilePath = Join-Path $strSubfolderPath 'FinalBurn Neo (ClrMame Pro XML, Arcade only).dat'

$strOutputFilePathMachineSummary = Join-Path '.' 'FBNeo_Arcade_DAT.csv'
$strOutputFilePathROMFileCRCs = Join-Path '.' 'FBNeo_Arcade_DAT_ROM_File_CRCs.csv'

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

function Get-ROMHashInfoRecursively {
    # This functions supports recursive ROM lookups in a MAME/FBNeo DAT to gather the ROM file
    # hashes in this machine and any parents. CHDs (disks) are not included
    #
    # The function takes four positional arguments.
    #
    # The first argument is a reference to an hashtable. Before calling the function,
    # a variable must be initialized to an empty hashtable (@{}). After completion of the
    # function, the hashtable is set to a series of keys, each of which is the CRC of a ROM
    # file in the specified ROM package. The value (in the key-value pair) is left $null
    #
    # The second argument is a string containing the short name of the machine (ROM package).
    #
    # The third argument is a reference to a hashtable of all the ROM information obtained
    # from the DAT, indexed by the ROM name.
    #
    # The function returns $true if successful, $false otherwise
    #
    # Example:
    # $strROMName = 'mario'
    # $hashtableROMCRCs = @{}
    # $boolSuccess = Get-ROMHashInfoRecursively ([ref]$hashtableROMCRCs) $strROMName ([ref]$hashtableEmulatorDAT)

    $refHashtableROMCRCs = $args[0]
    $strThisROMName = $args[1]
    $refHashtableDAT = $args[2]

    $strThisFunctionVersionNumber = [version]'1.0.20211230.0'

    $game = ($refHashtableDAT.Value).Item($strThisROMName)
    $boolParentSuccess = $true
    if ($null -ne $game.romof) {
        # This game has a parent ROM
        $boolParentSuccess = Get-ROMHashInfoRecursively $refHashtableROMCRCs ($game.romof) $refHashtableDAT
    }

    if ($boolParentSuccess -eq $false) {
        $false
    } else {
        $boolSuccess = $true

        if ($null -ne $game.rom) {
            @($game.rom) | ForEach-Object {
                $file = $_
                if ($file.status -ne 'nodump') {
                    if ($null -ne $file.crc) {
                        if (($refHashtableROMCRCs.Value).ContainsKey($file.crc) -eq $false) {
                            ($refHashtableROMCRCs.Value).Add($file.crc, $null)
                        }
                    }
                }
            }
        }

        $boolSuccess
    }
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

$VerbosePreference = $actionPreferenceNewVerbose

# Get the FBNeo DAT
$arrCommands = @(Get-Command Invoke-WebRequest)
$boolInvokeWebRequestAvailable = ($arrCommands.Count -ge 1)
if ($null -eq $strLocalXMLFilePath -and $boolInvokeWebRequestAvailable) {
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
        $arrNodes = @($HtmlNodeDownloadPage.SelectNodes('//a[@href]') | Where-Object { $_.InnerText.ToLower() -eq 'dats' })
        if ($arrNodes.Count -eq 0) {
            Write-Error ('Failed to download the FinalBurn Neo DAT file. Please download the file that looks like FinalBurn Neo (ClrMame Pro XML, Arcade only).dat from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalXMLFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalXMLFilePath to point to the path of the downloaded XML file.')
            break
        }
        $strNextURL = $arrNodes[0].Attributes['href'].Value

        $strURLBase = $strNextDownloadPageURL
        $strURLRelative = $strNextURL
        $strNextURL = Get-AbsoluteURLFromRelative $strURLBase $strURLRelative

        $strNextDownloadPageURL = $strNextURL
        $HtmlNodeDownloadPage = ConvertFrom-Html -URI $strNextDownloadPageURL
        $arrNodes = @($HtmlNodeDownloadPage.SelectNodes('//a[@href]') | Where-Object { $_.InnerText.ToLower().Contains('arcade only') })
        if ($arrNodes.Count -eq 0) {
            Write-Error ('Failed to download the FinalBurn Neo DAT file. Please download the file that looks like FinalBurn Neo (ClrMame Pro XML, Arcade only).dat from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalXMLFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalXMLFilePath to point to the path of the downloaded XML file.')
            break
        }
        $strNextURL = $arrNodes[0].Attributes['href'].Value

        $strURLBase = $strNextDownloadPageURL
        $strURLRelative = $strNextURL
        $strNextURL = Get-AbsoluteURLFromRelative $strURLBase $strURLRelative

        $strNextDownloadPageURL = $strNextURL
        $HtmlNodeDownloadPage = ConvertFrom-Html -URI $strNextDownloadPageURL
        $arrNodes = @($HtmlNodeDownloadPage.SelectNodes('//a[@href]') | Where-Object { $_.InnerText.ToLower() -like '*download*' })
        if ($arrNodes.Count -eq 0) {
            Write-Error ('Failed to download the FinalBurn Neo DAT file. Please download the file that looks like FinalBurn Neo (ClrMame Pro XML, Arcade only).dat from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalXMLFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalXMLFilePath to point to the path of the downloaded XML file.')
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
    Write-Verbose ('Downloading DAT from ' + $strEffectiveURL + '...')
    $VerbosePreference = $actionPreferenceFormerVerbose
    Invoke-WebRequest -Uri $strEffectiveURL -OutFile (Join-Path $strSubfolderPath 'FinalBurn Neo (ClrMame Pro XML, Arcade only).dat')
    $VerbosePreference = $actionPreferenceNewVerbose

    if (Test-Path (Join-Path $strSubfolderPath 'FinalBurn Neo (ClrMame Pro XML, Arcade only).dat')) {
        # Successful download
        $strAbsoluteXMLFilePath = (Resolve-Path (Join-Path $strSubfolderPath 'FinalBurn Neo (ClrMame Pro XML, Arcade only).dat')).Path
        Write-Verbose ('Loading DAT into memory and converting it to XML object...')
        $strContent = [System.IO.File]::ReadAllText($strAbsoluteXMLFilePath)
    } else {
        Write-Error ('Failed to download the FinalBurn Neo DAT file. Please download the file that looks like FinalBurn Neo (ClrMame Pro XML, Arcade only).dat from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalXMLFilePath + "`n`n" + 'Once downloaded, set the script variable $strLocalXMLFilePath to point to the path of the downloaded XML file.')
        break
    }
} else {
    if ((Test-Path $strLocalXMLFilePath) -ne $true) {
        Write-Error ('The FinalBurn Neo DAT file is missing. Please download the file that looks like FinalBurn Neo (ClrMame Pro XML, Arcade only).dat from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strDownloadPageURL + "`n`n" + 'File Location:' + "`n" + $strLocalXMLFilePath)
        break
    }
    $strAbsoluteXMLFilePath = (Resolve-Path $strLocalXMLFilePath).Path
    Write-Verbose ('Loading DAT into memory and converting it to XML object...')
    $strContent = [System.IO.File]::ReadAllText($strAbsoluteXMLFilePath)
}

# Convert it to XML
$xmlFBNeo = [xml]$strContent

Write-Verbose ('Creating a hashtable of ROM package information for rapid lookup by name...')
$hashtableFBNeo = New-BackwardCompatibleCaseInsensitiveHashtable
@($xmlFBNeo.datafile.game) | ForEach-Object {
    $game = $_
    $hashtableFBNeo.Add($game.name, $game)
}

Write-Verbose ('Processing ROM packages...')

$intTotalROMPackages = @($xmlFBNeo.datafile.game).Count
$intCurrentROMPackage = 1
$timeDateStartOfProcessing = Get-Date

$arrCSVFBNeo = @()
$arrCSVFBNeoROMCRCs = @()

@($xmlFBNeo.datafile.game) | ForEach-Object {
    $game = $_

    if ($intCurrentROMPackage -ge 101) {
        $timeDateCurrent = Get-Date
        $timeSpanElapsed = $timeDateCurrent - $timeDateStartOfProcessing
        $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / ($intCurrentROMPackage - 1) * $intTotalROMPackages
        $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
        $doublePercentComplete = ($intCurrentROMPackage - 1) / $intTotalROMPackages * 100
        Write-Progress -Activity 'Processing FBNeo ROM Packages' -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
    }

    $PSCustomObjectMachineSummary = New-Object PSCustomObject
    $PSCustomObjectROMFileCRCInfo = New-Object PSCustomObject

    $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'ROMName' -Value $game.name
    $PSCustomObjectROMFileCRCInfo | Add-Member -MemberType NoteProperty -Name 'ROMName' -Value $game.name

    $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMName' -Value $game.name
    $PSCustomObjectROMFileCRCInfo | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMName' -Value $game.name

    if ($null -eq $game.description) {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMDisplayName' -Value ''
        $PSCustomObjectROMFileCRCInfo | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMDisplayName' -Value ''
    } else {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMDisplayName' -Value $game.description
        $PSCustomObjectROMFileCRCInfo | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMDisplayName' -Value $game.description
    }

    ###########################################################################################

    if ($null -eq $game.manufacturer) {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_Manufacturer' -Value ''
    } else {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_Manufacturer' -Value $game.manufacturer
    }
    if ($null -eq $game.year) {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_Year' -Value ''
    } else {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_Year' -Value $game.year
    }
    if ($null -eq $game.cloneof) {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_CloneOf' -Value ''
    } else {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_CloneOf' -Value $game.cloneof
    }

    if (($null -eq $game.isbios) -or ($game.isbios -eq 'no')) {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_IsBIOSROM' -Value 'False'
    } else {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_IsBIOSROM' -Value 'True'
    }

    $boolROMPackageContainsROMs = $false
    $boolROMPackageContainsCHD = $false
    $boolROMFunctional = Test-MachineCompletelyFunctionalRecursively ([ref]$boolROMPackageContainsROMs) ([ref]$boolROMPackageContainsCHD) ($game.name) ([ref]$hashtableFBNeo)

    if ($boolROMFunctional -eq $true) {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_FunctionalROMPackage' -Value 'True'
    } else {
        $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_FunctionalROMPackage' -Value 'False'
    }

    ###########################################################################################

    $hashtableROMFileCRCs = New-BackwardCompatibleCaseInsensitiveHashtable
    $boolSuccess = Get-ROMHashInfoRecursively ([ref]$hashtableROMFileCRCs) ($game.name) ([ref]$hashtableFBNeo)

    if ($boolSuccess -eq $true) {
        $arrSortedCRCs = @(@($hashtableROMFileCRCs.Keys) | Sort-Object)
        if ($arrSortedCRCs.Count -ge 1) {
            $strSortedCRCs = [string]::Join("`t", $arrSortedCRCs)
        } else {
            $strSortedCRCs = ''
        }
    } else {
        $strSortedCRCs = ''
    }
    $PSCustomObjectROMFileCRCInfo | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMFileString' -Value $strSortedCRCs

    ###########################################################################################

    $strOverallStatus = 'Unknown'
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
        }
    }

    $PSCustomObjectMachineSummary | Add-Member -MemberType NoteProperty -Name 'FBNeo_OverallStatus' -Value $strOverallStatus

    $arrCSVFBNeo += $PSCustomObjectMachineSummary
    $arrCSVFBNeoROMCRCs += $PSCustomObjectROMFileCRCInfo

    $intCurrentROMPackage++
}

Write-Verbose ('Exporting results to CSV: ' + $strOutputFilePathMachineSummary)
$arrCSVFBNeo | Sort-Object -Property @('ROMName') |
    Export-Csv -Path $strOutputFilePathMachineSummary -NoTypeInformation

$arrCSVFBNeoROMCRCs | Sort-Object -Property @('ROMName') |
    Export-Csv -Path $strOutputFilePathROMFileCRCs -NoTypeInformation

$VerbosePreference = $actionPreferenceFormerVerbose
