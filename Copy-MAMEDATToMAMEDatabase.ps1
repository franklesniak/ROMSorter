# Copy-MAMEDATToMAMEDatabase.ps1

$strThisScriptVersionNumber = [version]'1.0.20211228.0'

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
$actionPreferenceNewDebug = $DebugPreference
$actionPreferenceFormerDebug = $DebugPreference
$strLocalMAMEDATCSV = $null
$strCSVOutputFile = $null

#region Inputs
###############################################################################################
# This script requires the current version of MAME's DAT converted to CSV. Set the following
# path to point to the corresponding CSV:
$strLocalMAMEDATCSV = Join-Path '.' 'MAME_DAT.csv'

# This script makes a copy of the MAME_DAT.csv file and stores it in the following file name:
$strCSVOutputFile = Join-Path '.' 'MAME_Database.csv'

# Comment-out the following line if you prefer that the script operate silently.
$actionPreferenceNewVerbose = [System.Management.Automation.ActionPreference]::Continue

# Remove the comment from the following line if you prefer that the script output extra
# debugging information.
# $actionPreferenceNewDebug = [System.Management.Automation.ActionPreference]::Continue

###############################################################################################
#endregion Inputs

$VerbosePreference = $actionPreferenceNewVerbose
$DebugPreference = $actionPreferenceNewDebug

$boolErrorOccurred = $false

if ((Test-Path $strLocalMAMEDATCSV) -ne $true) {
    Write-Error ('The input file "' + $strLocalMAMEDATCSV + '" is missing. Please generate it using the "Convert-MAMEDATToCsv" script and then re-run this script')
    $boolErrorOccurred = $true
}

if ($boolErrorOccurred -eq $true) {
    break
}

Write-Verbose "Making a copy of the MAME DAT..."
Copy-Item -Path $strLocalMAMEDATCSV -Destination $strCSVOutputFile -Force
Write-Verbose "Done!"

$VerbosePreference = $actionPreferenceFormerVerbose
$DebugPreference = $actionPreferenceFormerDebug
