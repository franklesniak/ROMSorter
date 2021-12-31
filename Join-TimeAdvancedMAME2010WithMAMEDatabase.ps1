# Join-TimeAdvancedMAME2010WithMAMEDatabase.ps1

$strThisScriptVersionNumber = [version]'1.0.20211231.0'

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

$actionPreferenceFormerVerbose = $VerbosePreference
$actionPreferenceFormerDebug = $DebugPreference

$strLocalMAMEDatabaseCSV = $null
$boolBackupMAMEDatabaseBeforeOverwrite = $true
$strLocalMAMEDatabaseBackupPrefix = $null
$strLocalMAMEDatabaseBackupSuffix = $null
$strScriptToGenerateTimeAdvancedDAT = 'coresponding "Convert-..."'
$strLocalTimeAdvancedDATToJoinCSV = $null
$strTimeAdvancedDATToJoinDisplayName = 'this time-advanced DAT'
$actionPreferenceNewVerbose = $VerbosePreference
$actionPreferenceNewDebug = $DebugPreference

#region Inputs
###############################################################################################
# This script requires the current version of MAME's DAT converted to CSV. Optionally, the DAT
# may have already been joined with DATs from other versions of MAME or other emulators. Set
# the following path to point to the corresponding CSV:
$strLocalMAMEDatabaseCSV = Join-Path '.' 'MAME_Database.csv'

# This script will make a backup copy of the previous file before overwriting it with the new
# results. If you don't want to take a backup, uncomment the following line:
# $boolBackupMAMEDatabaseBeforeOverwrite = $false

# Change this if you want to effect the name of the backup file:
$strLocalMAMEDatabaseBackupPrefix = Join-Path '.' 'MAME_Database_Backup_'
$strLocalMAMEDatabaseBackupSuffix = '.csv'

# If the time-advanced CSV below is missing, the user is prompted to run the following script:
$strScriptToGenerateTimeAdvancedDAT = '"Convert-MAME2010CSVToRenamedROMPackagesMatchingCurrentMAME"'

# This script also requires a "time-advanced" version of another MAME DAT (i.e., the one to
# join to the database), converted to CSV format. Time advancement occurs by applying the
# Progetto Snaps RenameSet against the DAT:
$strLocalTimeAdvancedDATToJoinCSV = Join-Path '.' 'MAME_2010_DAT_With_Time-Advanced_ROM_Package_Names.csv'

# This display name is used in progress output:
$strTimeAdvancedDATToJoinDisplayName = 'the MAME 2010 time-advanced DAT'

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

$dateTimeStart = Get-Date

if ((Test-Path $strLocalMAMEDatabaseCSV) -ne $true) {
    Write-Error ('The input file "' + $strLocalMAMEDatabaseCSV + '" is missing. Please generate it using the "Copy-MAMEDATToMAMEDatabase" script and then re-run this script')
    $boolErrorOccurred = $true
}

if ((Test-Path $strLocalTimeAdvancedDATToJoinCSV) -ne $true) {
    Write-Error ('The input file "' + $strLocalTimeAdvancedDATToJoinCSV + '" is missing. Please generate it using the ' + $strScriptToGenerateTimeAdvancedDAT + ' script and then re-run this script')
    $boolErrorOccurred = $true
}

$VerbosePreference = $actionPreferenceFormerVerbose
$arrModules = @(Get-Module Join-Object -ListAvailable)
$VerbosePreference = $actionPreferenceNewVerbose

if ($arrModules.Count -eq 0) {
    Write-Error 'This script requires the module "Join-Object". On PowerShell version 5.0 and newer, it can be instaled using the command "Install-Module Join-Object". Please install it and re-run the script'
    $boolErrorOccurred = $true
}

if ($boolErrorOccurred -eq $true) {
    break
}

Write-Verbose ('Importing the MAME database...')
$arrMAMEDatabase = @(Import-Csv -Path $strLocalMAMEDatabaseCSV)

Write-Verbose ('Importing ' + $strTimeAdvancedDATToJoinDisplayName + '...')
$arrTimeAdvancedDATToJoin = @(Import-Csv -Path $strLocalTimeAdvancedDATToJoinCSV)

$arrLoadedModules = @(Get-Module Join-Object)
if ($arrLoadedModules.Count -eq 0) {
    $VerbosePreference = $actionPreferenceFormerVerbose
    Import-Module Join-Object
    $VerbosePreference = $actionPreferenceNewVerbose
}

Write-Verbose ('Joining the MAME database with ' + $strTimeAdvancedDATToJoinDisplayName + '. This may take a while...')
$arrJoinedData = Join-Object -Left $arrMAMEDatabase -Right $arrTimeAdvancedDATToJoin -LeftJoinProperty 'ROMName' -RightJoinProperty 'ROMName' -Type 'AllInBoth' -AddKey 'ROMName' -ExcludeLeftProperties 'ROMName' -ExcludeRightProperties 'ROMName'

if ($boolBackupMAMEDatabaseBeforeOverwrite -eq $true) {
    $strBackupFileName = ($strLocalMAMEDatabaseBackupPrefix + ([string]($dateTimeStart.Year) + '-' + [string]($dateTimeStart.Month) + '-' + [string]($dateTimeStart.Day) + '_' + [string]($dateTimeStart.Hour) + [string]($dateTimeStart.Minute) + [string]($dateTimeStart.Second)) + $strLocalMAMEDatabaseBackupSuffix)
    Copy-Item -Path $strLocalMAMEDatabaseCSV -Destination $strBackupFileName -Force
}

Write-Verbose ('Exporting results to CSV: ' + $strCSVOutputFile)
$arrJoinedData | Sort-Object -Property @('ROMName') | Export-Csv -Path $strLocalMAMEDatabaseCSV -NoTypeInformation

Write-Verbose "Done!"

$VerbosePreference = $actionPreferenceFormerVerbose
$DebugPreference = $actionPreferenceFormerDebug
