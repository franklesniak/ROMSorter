# Convert-FBNeoCSVToRenamedROMPackagesMatchingCurrentMAME.ps1

$strThisScriptVersionNumber = [version]'1.0.20220522.0'

#region License
#######################################################################################
# Copyright 2022 Frank Lesniak

# Permission is hereby granted, free of charge, to any person obtaining a copy of this
# software and associated documentation files (the "Software"), to deal in the Software
# without restriction, including without limitation the rights to use, copy, modify,
# merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
# permit persons to whom the Software is furnished to do so, subject to the following
# conditions:

# The above copyright notice and this permission notice shall be included in all copies
# or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
# INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
# PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
# HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
# CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
# OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#######################################################################################
#endregion License

#region DownloadLocationNotice
# The most up-to-date version of this script can be found on the author's GitHub repository
# at https://github.com/franklesniak/ROMSorter
#endregion DownloadLocationNotice

$actionPreferenceNewVerbose = $VerbosePreference
$actionPreferenceFormerVerbose = $VerbosePreference
$actionPreferenceNewDebug = $DebugPreference
$actionPreferenceFormerDebug = $DebugPreference

#region Inputs
#######################################################################################
# This script requires the FBNeo DAT converted to CSV. Set the following path to point
# to the corresponding CSV:
$strPathToFBNeoROMPackageMetadataCSV = Join-Path '.' 'FBNeo_Arcade_DAT.csv'

# This script also requires the MAME DAT converted to CSV. Set the following path to
# point to the corresponding CSV:
$strPathToMAMEROMPackageMetadataCSV = Join-Path '.' 'MAME_DAT.csv'

# The file will be processed and output as a CSV to
# .\Progetto_Snaps_RenameSet.csv
# or if on Linux / MacOS: ./Progetto_Snaps_RenameSet.csv
$strCSVOutputFile = Join-Path '.' 'FBNeo_DAT_With_ROM_Package_Names_String-Matched_To_MAME.csv'

# Comment-out the following line if you prefer that the script operate silently.
$actionPreferenceNewVerbose = [System.Management.Automation.ActionPreference]::Continue

# Remove the comment from the following line if you prefer that the script output extra
# debugging information.
# $actionPreferenceNewDebug = [System.Management.Automation.ActionPreference]::Continue
#######################################################################################
#endregion Inputs

$VerbosePreference = $actionPreferenceNewVerbose
$DebugPreference = $actionPreferenceNewDebug

$boolErrorOccurred = $false

if ((Test-Path $strPathToFBNeoROMPackageMetadataCSV) -ne $true) {
    Write-Error ('The input file "' + $strPathToFBNeoROMPackageMetadataCSV + '" is missing. Please generate it using the corresponding "Convert-..." script and then re-run this script')
    $boolErrorOccurred = $true
}

if ((Test-Path $strPathToMAMEROMPackageMetadataCSV) -ne $true) {
    Write-Error ('The input file "' + $strPathToMAMEROMPackageMetadataCSV + '" is missing. Please generate it using the corresponding "Convert-..." script and then re-run this script')
    $boolErrorOccurred = $true
}

if ($boolErrorOccurred -eq $true) {
    break
}

Write-Verbose "Loading FBNeo Arcade ROM set's ROM package metadata..."
$arrFBNeoROMPackageMetadata = Import-Csv $strPathToFBNeoROMPackageMetadataCSV
$hashtableFBNeoROMNameToROMInfo = @{}
$arrFBNeoROMPackageMetadata | ForEach-Object {
    $hashtableFBNeoROMNameToROMInfo.Add($_.FBNeo_ROMName, $_)
}

Write-Verbose "Loading MAME ROM set's ROM package metadata..."
$arrMAMEROMPackageMetadata = Import-Csv $strPathToMAMEROMPackageMetadataCSV
$hashtableMAMEROMNameToROMInfo = @{}
$arrMAMEROMPackageMetadata | ForEach-Object {
    $hashtableMAMEROMNameToROMInfo.Add($_.MAME_ROMName, $_)
}

#region Initialize hashtables for holding matches
Write-Verbose 'Initializing data structures for holding string-matching metadata...'
$hashtableFBNeoROMNameToMatches = @{}

$hashtableMAMEROMNameToAllMatches = @{}
$arrMAMEROMPackageMetadata | ForEach-Object {
    $arrayList = New-Object -TypeName 'System.Collections.ArrayList'
    $hashtableMAMEROMNameToAllMatches.Add($_.MAME_ROMName, $arrayList)
}
#endregion Initialize hashtables for holding matches

$intTotalToProcess = ($arrFBNeoROMPackageMetadata.Count) * ($arrMAMEROMPackageMetadata.Count)
$intCurrentItem = 0
$refIntCurrentItem = [ref]$intCurrentItem
$timeDateStartOfProcessing = Get-Date

$arrFBNeoROMPackageMetadata | ForEach-Object {
    $strFBNeoROMName = $_.FBNeo_ROMName
    $strFBNeoROMDisplayName = $_.FBNeo_ROMDisplayName

    $job = $arrMAMEROMPackageMetadata | ForEach-Object -AsJob -ThrottleLimit 16 -Parallel {
        $strMAMEROMName = $_.MAME_ROMName
        $strMAMEROMDisplayName = $_.MAME_ROMDisplayName
        $dblJaccardIndexROMName = Get-JaccardIndex -a $using:strFBNeoROMName -b $strMAMEROMName -CaseSensitive:$false
        $dblJaccardIndexDisplayName = Get-JaccardIndex -a $using:strFBNeoROMDisplayName -b $strMAMEROMDisplayName -CaseSensitive:$false
        $dblAvgScore = ($dblJaccardIndexROMName + $dblJaccardIndexDisplayName) / 2

        $PSObjectMatchToMAME = New-Object PSObject
        $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'AverageScore' -Value $dblAvgScore
        $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToFBNeoROMName' -Value $dblJaccardIndexROMName
        $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMName' -Value $strFBNeoROMName
        $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToFBNeoROMDisplayName' -Value $dblJaccardIndexDisplayName
        $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMDisplayName' -Value $strFBNeoROMDisplayName

        # Stash this match info on the corresponding MAME hashtable:
        (($using:hashtableMAMEROMNameToAllMatches).Item($strMAMEROMName)).Add($PSObjectMatchToMAME)

        $PSObjectMatchToFBNeo = New-Object PSObject
        $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'AverageScore' -Value $dblAvgScore
        $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToMAMEROMName' -Value $dblJaccardIndexROMName
        $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'MAME_ROMName' -Value $strMAMEROMName
        $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToMAMEROMDisplayName' -Value $dblJaccardIndexDisplayName
        $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'MAME_ROMDisplayName' -Value $strMAMEROMDisplayName

        $null = [Threading.Interlocked]::Increment($using:refIntCurrentItem)

        return $PSObjectMatchToFBNeo
    }

    #While $job is running, update progress bar
    while ($job.State -eq 'Running') {
        if ($intCurrentItem -ge 100) {
            $timeDateCurrent = Get-Date
            $timeSpanElapsed = $timeDateCurrent - $timeDateStartOfProcessing
            $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / $intCurrentItem * $intTotalToProcess
            $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
            $doublePercentComplete = $intCurrentItem / $intTotalToProcess * 100
            Write-Progress -Activity 'Comparing FBNeo ROMs to MAME ROMs' -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
        }
        Start-Sleep -Milliseconds 500
    }

    $arrMatches = Receive-Job $job |
        Sort-Object -Property 'AverageScore' -Descending |
        Select-Object -First 20

    $hashtableFBNeoROMNameToMatches.Add($strFBNeoROMName, $arrMatches)
}

#     $arrMatches = $arrMAMEROMPackageMetadata | ForEach-Object -ThrottleLimit 5 -Parallel {
#         if ($using:intCurrentItem -ge 100) {
#             $timeDateCurrent = Get-Date
#             $timeSpanElapsed = $timeDateCurrent - $using:timeDateStartOfProcessing
#             $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / $using:intCurrentItem * $intTotalToProcess
#             $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
#             $doublePercentComplete = $using:intCurrentItem / $using:intTotalToProcess * 100
#             Write-Progress -Activity 'Comparing FBNeo ROMs to MAME ROMs' -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
#         }
#         $strMAMEROMName = $_.MAME_ROMName
#         $strMAMEROMDisplayName = $_.MAME_ROMDisplayName
#         $dblJaccardIndexROMName = Get-JaccardIndex -a $using:strFBNeoROMName -b $strMAMEROMName -CaseSensitive:$false
#         $dblJaccardIndexDisplayName = Get-JaccardIndex -a $using:strFBNeoROMDisplayName -b $strMAMEROMDisplayName -CaseSensitive:$false
#         $dblAvgScore = ($dblJaccardIndexROMName + $dblJaccardIndexDisplayName) / 2

#         $PSObjectMatchToMAME = New-Object PSObject
#         $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'AverageScore' -Value $dblAvgScore
#         $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToFBNeoROMName' -Value $dblJaccardIndexROMName
#         $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMName' -Value $strFBNeoROMName
#         $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToFBNeoROMDisplayName' -Value $dblJaccardIndexDisplayName
#         $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMDisplayName' -Value $strFBNeoROMDisplayName

#         # Stash this match info on the corresponding MAME hashtable:
#         (($using:hashtableMAMEROMNameToAllMatches).Item($strMAMEROMName)).Add($PSObjectMatchToMAME)

#         $PSObjectMatchToFBNeo = New-Object PSObject
#         $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'AverageScore' -Value $dblAvgScore
#         $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToMAMEROMName' -Value $dblJaccardIndexROMName
#         $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'MAME_ROMName' -Value $strMAMEROMName
#         $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToMAMEROMDisplayName' -Value $dblJaccardIndexDisplayName
#         $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'MAME_ROMDisplayName' -Value $strMAMEROMDisplayName

#         $null = [Threading.Interlocked]::Increment($using:refIntCurrentItem)

#         return $PSObjectMatchToFBNeo
#     } | Sort-Object -Property 'AverageScore' -Descending | Select-Object -First 20
#     $hashtableFBNeoROMNameToMatches.Add($strFBNeoROMName, $arrMatches)
# }

$VerbosePreference = $actionPreferenceFormerVerbose
$DebugPreference = $actionPreferenceFormerDebug
