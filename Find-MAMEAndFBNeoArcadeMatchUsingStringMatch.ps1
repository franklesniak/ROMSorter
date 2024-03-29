# Find-MAMEAndFBNeoArcadeMatchUsingStringMatch.ps1

$strThisScriptVersionNumber = [version]'0.3.20220616.0'

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

Write-Verbose "Loading MAME ROM set's ROM package metadata..."
$arrMAMEROMPackageMetadata = Import-Csv $strPathToMAMEROMPackageMetadataCSV

#region Perform preprocessing on FBNeo ROM package metadata
# This speeds up processing later by eliminating the need to convert to a lowercase
# character array over and over...
Write-Verbose 'Performing pre-processing on FBNeo ROM package metadata...'
$hashtableFBNeoROMNameToROMInfo = @{}
$arrFBNeoROMPackageMetadata | ForEach-Object {
    $strLowercaseROMName = ($_.FBNeo_ROMName).ToLower()
    $strLowercaseROMDisplayName = ($_.FBNeo_ROMDisplayName).ToLower()
    $arrCharLowercaseROMName = @($strLowercaseROMName.ToCharArray())
    $arrCharLowercaseROMDisplayName = @($strLowercaseROMDisplayName.ToCharArray())
    $hashsetLowercaseROMName = New-Object System.Collections.Generic.HashSet[System.Char] (, ($arrCharLowercaseROMName -as 'System.Char[]'))
    $hashsetLowercaseROMDisplayName = New-Object System.Collections.Generic.HashSet[System.Char] (, ($arrCharLowercaseROMDisplayName -as 'System.Char[]'))
    $_ | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMName_Lowercase_Hashset' -Value $hashsetLowercaseROMName
    $_ | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMDisplayName_Lowercase_Hashset' -Value $hashsetLowercaseROMDisplayName
    $hashtableFBNeoROMNameToROMInfo.Add($_.FBNeo_ROMName, $_)
}
#endregionPerform preprocessing on FBNeo ROM package metadata

#region Perform preprocessing on MAME ROM package metadata
# This speeds up processing later by eliminating the need to convert to a lowercase
# character array over and over...
Write-Verbose 'Performing pre-processing on MAME ROM package metadata...'
$hashtableMAMEROMNameToROMInfo = @{}
$arrMAMEROMPackageMetadata | ForEach-Object {
    $strLowercaseROMName = ($_.MAME_ROMName).ToLower()
    $strLowercaseROMDisplayName = ($_.MAME_ROMDisplayName).ToLower()
    $arrCharLowercaseROMName = @($strLowercaseROMName.ToCharArray())
    $arrCharLowercaseROMDisplayName = @($strLowercaseROMDisplayName.ToCharArray())
    $hashsetLowercaseROMName = New-Object System.Collections.Generic.HashSet[System.Char] (, ($arrCharLowercaseROMName -as 'System.Char[]'))
    $hashsetLowercaseROMDisplayName = New-Object System.Collections.Generic.HashSet[System.Char] (, ($arrCharLowercaseROMDisplayName -as 'System.Char[]'))
    $_ | Add-Member -MemberType NoteProperty -Name 'MAME_ROMName_Lowercase_Hashset' -Value $hashsetLowercaseROMName
    $_ | Add-Member -MemberType NoteProperty -Name 'MAME_ROMDisplayName_Lowercase_Hashset' -Value $hashsetLowercaseROMDisplayName
    $hashtableMAMEROMNameToROMInfo.Add($_.MAME_ROMName, $_)
}
#endregion Perform preprocessing on MAME ROM package metadata

#region Initialize hashtables for holding matches
Write-Verbose 'Initializing data structures for holding string-matching metadata...'
$hashtableFBNeoROMNameToAllMatches = @{}
$arrFBNeoROMPackageMetadata | ForEach-Object {
    $arrayListMAMEMatches = New-Object -TypeName 'System.Collections.ArrayList'
    $hashtableFBNeoROMNameToAllMatches.Add($_.FBNeo_ROMName, $arrayListMAMEMatches)
}

$hashtableMAMEROMNameToAllMatches = @{}
$arrMAMEROMPackageMetadata | ForEach-Object {
    $arrayListFBNeoMatches = New-Object -TypeName 'System.Collections.ArrayList'
    $hashtableMAMEROMNameToAllMatches.Add($_.MAME_ROMName, $arrayListFBNeoMatches)
}
#endregion Initialize hashtables for holding matches

$intTotalToProcess = ($arrFBNeoROMPackageMetadata.Count) * ($arrMAMEROMPackageMetadata.Count)
$intCurrentItem = 0
$timeDateStartOfProcessing = Get-Date

$arrFBNeoROMPackageMetadata | ForEach-Object {
    $refStrFBNeoROMName = [ref]($_.FBNeo_ROMName)
    $refStrFBNeoROMDisplayName = [ref]($_.FBNeo_ROMDisplayName)
    $refHashsetFBNeoLowercaseROMName = [ref]($_.FBNeo_ROMName_Lowercase_Hashset)
    $refHashsetFBNeoLowercaseROMDisplayName = [ref]($_.FBNeo_ROMDisplayName_Lowercase_Hashset)

    $arrMAMEROMPackageMetadata | ForEach-Object {
        if ($intCurrentItem -ge 1000) {
            $timeDateCurrent = Get-Date
            $timeSpanElapsed = $timeDateCurrent - $timeDateStartOfProcessing
            $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / $intCurrentItem * $intTotalToProcess
            $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
            $doublePercentComplete = $intCurrentItem / $intTotalToProcess * 100
            Write-Progress -Activity 'Comparing FBNeo ROMs to MAME ROMs' -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
        }

        $refStrMAMEROMName = [ref]($_.MAME_ROMName)
        $refStrMAMEROMDisplayName = [ref]($_.MAME_ROMDisplayName)
        $refHashsetMAMELowercaseROMName = [ref]($_.MAME_ROMName_Lowercase_Hashset)
        $refHashsetMAMELowercaseROMDisplayName = [ref]($_.MAME_ROMDisplayName_Lowercase_Hashset)

        $hashsetWorkingROMNameIntersection = New-Object System.Collections.Generic.HashSet[System.Char] ($refHashsetMAMELowercaseROMName.Value)
        $hashsetWorkingROMDisplayNameIntersection = New-Object System.Collections.Generic.HashSet[System.Char] ($refHashsetMAMELowercaseROMDisplayName.Value)
        $hashsetWorkingROMNameUnion = New-Object System.Collections.Generic.HashSet[System.Char] ($refHashsetMAMELowercaseROMName.Value)
        $hashsetWorkingROMDisplayNameUnion = New-Object System.Collections.Generic.HashSet[System.Char] ($refHashsetMAMELowercaseROMDisplayName.Value)

        $hashsetWorkingROMNameIntersection.IntersectWith($refHashsetFBNeoLowercaseROMName.Value)
        $hashsetWorkingROMDisplayNameIntersection.IntersectWith($refHashsetFBNeoLowercaseROMDisplayName.Value)
        $hashsetWorkingROMNameUnion.UnionWith($refHashsetFBNeoLowercaseROMName.Value)
        $hashsetWorkingROMDisplayNameUnion.UnionWith($refHashsetFBNeoLowercaseROMDisplayName.Value)

        $dblJaccardIndexROMName = (($hashsetWorkingROMNameIntersection.Count) / ($hashsetWorkingROMNameUnion.Count))
        $dblJaccardIndexROMDisplayName = (($hashsetWorkingROMDisplayNameIntersection.Count) / ($hashsetWorkingROMDisplayNameUnion.Count))
        $dblAvgScore = ($dblJaccardIndexROMName + $dblJaccardIndexROMDisplayName) / 2

        if ($dblAvgScore -gt 0.5) {
            $PSObjectMatchToMAME = New-Object PSObject
            $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'AverageScore' -Value $dblAvgScore
            $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToFBNeoROMName' -Value $dblJaccardIndexROMName
            $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMName' -Value ($refStrFBNeoROMName.Value)
            $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToFBNeoROMDisplayName' -Value $dblJaccardIndexDisplayName
            $PSObjectMatchToMAME | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMDisplayName' -Value ($refStrFBNeoROMDisplayName.Value)

            # Stash this match info on the corresponding MAME hashtable:
            (($hashtableMAMEROMNameToAllMatches).Item($refStrMAMEROMName.Value)).Add($PSObjectMatchToMAME) | Out-Null

            $PSObjectMatchToFBNeo = New-Object PSObject
            $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'AverageScore' -Value $dblAvgScore
            $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToMAMEROMName' -Value $dblJaccardIndexROMName
            $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'MAME_ROMName' -Value ($refStrMAMEROMName.Value)
            $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'JaccardIndexToMAMEROMDisplayName' -Value $dblJaccardIndexDisplayName
            $PSObjectMatchToFBNeo | Add-Member -MemberType NoteProperty -Name 'MAME_ROMDisplayName' -Value ($refStrMAMEROMDisplayName.Value)

            # Stash this match info on the corresponding FBNeo hashtable:
            (($hashtableFBNeoROMNameToAllMatches).Item($refStrFBNeoROMName.Value)).Add($PSObjectMatchToFBNeo) | Out-Null
        }

        $intCurrentItem++
    } # | Sort-Object -Property 'AverageScore' -Descending | Select-Object -First 20
}

$arrayListFBNeoToMAMEExactMatches = New-Object -TypeName 'System.Collections.ArrayList'
$arrayListFBNeoToMAMEMatchesWithDifferentROMNames = New-Object -TypeName 'System.Collections.ArrayList'
$arrayListUnmatchedFBNeoROMs = New-Object -TypeName 'System.Collections.ArrayList'

$intTotalToProcess = @($hashtableFBNeoROMNameToAllMatches.Keys).Count
$intCurrentItem = 0
$timeDateStartOfProcessing = Get-Date

$hashtableFBNeoROMNameToAllMatches.Keys | Sort-Object | ForEach-Object {
    if (($intCurrentItem / $intTotalToProcess) -ge 0.02) {
        $timeDateCurrent = Get-Date
        $timeSpanElapsed = $timeDateCurrent - $timeDateStartOfProcessing
        $doubleTotalProcessingTimeInSeconds = $timeSpanElapsed.TotalSeconds / $intCurrentItem * $intTotalToProcess
        $doubleRemainingProcessingTimeInSeconds = $doubleTotalProcessingTimeInSeconds - $timeSpanElapsed.TotalSeconds
        $doublePercentComplete = $intCurrentItem / $intTotalToProcess * 100
        Write-Progress -Activity 'Confirming FBNeo to MAME Matches' -PercentComplete $doublePercentComplete -SecondsRemaining $doubleRemainingProcessingTimeInSeconds
    }

    $strFBNeoROMName = $_
    $arrThisFBNeoROMMAMEMatches = @($hashtableFBNeoROMNameToAllMatches.Item($strFBNeoROMName) | Sort-Object -Property 'AverageScore' -Descending)
    if ($arrThisFBNeoROMMAMEMatches.Count -ge 1) {
        $strBestMatchingMAMEROMName = ($arrThisFBNeoROMMAMEMatches[0]).MAME_ROMName
        $arrThisFBNeoROMsBestMAMEMatchMatches = @($hashtableMAMEROMNameToAllMatches.Item($strBestMatchingMAMEROMName) | Sort-Object -Property 'AverageScore' -Descending)
        if ($arrThisFBNeoROMsBestMAMEMatchMatches.Count -ge 1) {
            $strBestMatchingMAMEROMsBestMatchingFBNeoROMName = ($arrThisFBNeoROMsBestMAMEMatchMatches[0]).FBNeo_ROMName
            if ($strFBNeoROMName -eq $strBestMatchingMAMEROMName -and $strBestMatchingMAMEROMName -eq $strBestMatchingMAMEROMsBestMatchingFBNeoROMName) {
                # FBNeo-ROMName ---matched---> Same MAME_ROMName ---matched--->
                # FBNeo-ROMName
                $arrayListFBNeoToMAMEExactMatches.Add($strFBNeoROMName) | Out-Null
            } elseif ($strFBNeoROMName -eq $strBestMatchingMAMEROMsBestMatchingFBNeoROMName) {
                # FBNeo-ROMName ---matched---> Different MAME_ROMName ---matched--->
                # FBNeo-ROMName
                $PSObjectMatchedWithDifferentROMNames = New-Object PSObject
                $PSObjectMatchedWithDifferentROMNames | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMName' -Value $strFBNeoROMName
                $PSObjectMatchedWithDifferentROMNames | Add-Member -MemberType NoteProperty -Name 'MAME_ROMName' -Value $strBestMatchingMAMEROMName
                $arrayListFBNeoToMAMEMatchesWithDifferentROMNames.Add($PSObjectMatchedWithDifferentROMNames) | Out-Null
            } else {
                Write-Output ('Did not match ' + $strFBNeoROMName)
            }
        } else {
            Write-Warning ('The MAME ROM ' + $strBestMatchingMAMEROMName + ' was matched to FBNeo ROM ' + $strFBNeoROMName + '; however, the MAME ROM ' + $strBestMatchingMAMEROMName + ' had no matches.')
            $arrayListUnmatchedFBNeoROMs.Add($strFBNeoROMName) | Out-Null
        }
    } else {
        Write-Warning ('The FBNeo ROM ' + $strFBNeoROMName + ' had no matches.')
        $arrayListUnmatchedFBNeoROMs.Add($strFBNeoROMName) | Out-Null
    }
    $intCurrentItem++
}

$arrayListOutput = New-Object -TypeName 'System.Collections.ArrayList'

$arrayListFBNeoToMAMEExactMatches | ForEach-Object {
    $strThisFBNeoROMName = $_
    $PSObjectOutputMatch = New-Object PSObject
    $PSObjectOutputMatch | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMName' -Value $strThisFBNeoROMName
    $PSObjectOutputMatch | Add-Member -MemberType NoteProperty -Name 'MAME_ROMName' -Value $strThisFBNeoROMName
    $arrayListOutput.Add($PSObjectOutputMatch) | Out-Null
}

$arrayListFBNeoToMAMEMatchesWithDifferentROMNames | ForEach-Object {
    $PSObjectOutput = $_
    $arrayListOutput.Add($PSObjectOutput) | Out-Null
}

$arrayListUnmatchedFBNeoROMs | ForEach-Object {
    $strThisFBNeoROMName = $_
    $PSObjectOutputMatch = New-Object PSObject
    $PSObjectOutputMatch | Add-Member -MemberType NoteProperty -Name 'FBNeo_ROMName' -Value $strThisFBNeoROMName
    $PSObjectOutputMatch | Add-Member -MemberType NoteProperty -Name 'MAME_ROMName' -Value '????'
    $arrayListOutput.Add($PSObjectOutputMatch) | Out-Null
}

$arrayListOutput | Sort-Object -Property 'FBNeo_ROMName' | Export-Csv -Path $strCSVOutputFile -NoTypeInformation

$VerbosePreference = $actionPreferenceFormerVerbose
$DebugPreference = $actionPreferenceFormerDebug
