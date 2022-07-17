# ROMSorter
A scripted, repeatable, customizable approach to sorting ROMs for video game emulation.

## Arcade Emulation and ROM-Sorting
Please read the [Arcade Emulator Background Info page](./ARCADE_EMULATOR_BACKGROUND_INFO.md) before getting started; it includes required background information.
If you are analyzing or sorting ROM packages, or if you are building a ROM list for an arcade system, the process will not make much sense unless you read the background page.

## Methodology

### Convert MAME and FinalBurn Neo ROM DATs to Tabular (CSV) Format

1. Create a new, empty folder.
In PowerShell, switch the working directory to the new folder (`cd C:\ROMSorter`).
Set the script execution policy to bypass (or equivalent, e.g.: `Set-ExecutionPolicy Bypass -Scope Process`).
1. Run `Convert-MAMEDATToCSV.ps1` to automatically download the newest-released MAME DAT XML file and convert it to a tabular (CSV) data format. The DAT is downloaded from [The MAME team's website](https://www.mamedev.org/release.html), and the output file of this script is `MAME_DAT.csv`.
1. Run `Convert-MAME2003PlusDATToCSV.ps1` to automatically download the libretro team's cloned version of MAME 2003 Plus's DAT XML file and convert it to a tabular (CSV) data format. The DAT is downloaded from [GitHub](https://github.com/libretro/mame2003-plus-libretro), and the output file of this script is `MAME_2003_Plus_DAT.csv`.
1. Run `Convert-MAME2010DATToCSV.ps1` to automatically download the libretro team's cloned version of MAME 2010's DAT XML file and convert it to a tabular (CSV) data format. The DAT is downloaded from [GitHub](https://github.com/libretro/mame2010-libretro), and the output file of this script is `MAME_2010_DAT.csv`.
1. Run `Convert-FBNeoArcadeDATToCSV.ps1` to automatically download the libretro team's cloned version of FinalBurn Neo's (FBNeo's) DAT XML file and convert it to a tabular (CSV) data format. The DAT is downloaded from [GitHub](https://github.com/libretro/FBNeo), and the output file of the script is `FBNeo_Arcade_DAT.csv`.
