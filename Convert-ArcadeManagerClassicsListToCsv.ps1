# Convert-ArcadeManagerClassicsListToCsv.ps1 is simply takes the semicolon-separated values
# file obtained from:
# https://raw.githubusercontent.com/cosmo0/arcade-manager-data/master/csv/best/classics-all.csv
# and converts it to a proper CSV format for easier downstream handling

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

#region Inputs
# Download the classics-all.csv file from
# https://raw.githubusercontent.com/cosmo0/arcade-manager-data/master/csv/best/classics-all.csv
# and put it in the following folder:
# .\Arcade_Manager_Resources
# or if on Linux / MacOS: ./Arcade_Manager_Resources
# i.e., the folder that this script is in should have a subfolder called:
# Arcade_Manager_Resources
$strSubfolderPath = Join-Path '.' 'Arcade_Manager_Resources'

# The file will be processed and output as a CSV to
# .\Arcade_Manager_Classics_List.csv
# or if on Linux / MacOS: ./Arcade_Manager_Classics_List.csv
$strCSVOutputFile = Join-Path '.' 'Arcade_Manager_Classics_List.csv'

###############################################################################################
#endregion Inputs

$boolErrorOccurred = $false

# Arcade Manager Classics (All) "CSV" file (really, it's a semicolon-separated file)
$strURLArcadeManagerClassicsAll = 'https://raw.githubusercontent.com/cosmo0/arcade-manager-data/master/csv/best/classics-all.csv'
$strFilePathArcadeManagerClassicsAllSemicolonSeparated = Join-Path $strSubfolderPath 'classics-all.csv'

if ((Test-Path $strFilePathArcadeManagerClassicsAllSemicolonSeparated) -ne $true) {
    Write-Error ('The Arcade Manager Classics (All) "CSV" file is missing. Please download it from the following URL and place it in the following location.' + "`n`n" + 'URL: ' + $strURLArcadeManagerClassicsAll + "`n`n" + 'File Location:' + "`n" + $strFilePathArcadeManagerClassicsAllSemicolonSeparated)
    $boolErrorOccurred = $true
}

if ($boolErrorOccurred -eq $false) {
    # We have all the files, let's do stuff

    # Import as semicolon-separated values "CSV"
    $csvCurrentRomList = Import-Csv $strFilePathArcadeManagerClassicsAllSemicolonSeparated -Delimiter ';'

    # Add another column to the CSV indicating that the ROM was listed as classic in Arcade
    # Manager. It's redundant, but it makes certain analysis easier.
    $csvCurrentRomList = $csvCurrentRomList | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name 'ArcadeManagerListedAsClassic' -Value 'True'
        $_
    }

    # Export as comma-separated values (true CSV)
    $csvCurrentRomList | Export-Csv $strCSVOutputFile -NoTypeInformation
}
