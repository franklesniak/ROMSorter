# EOL Normalization Review

Separate byte-only phase, reviewed 2026-11-11. The functional/template content changes were prepared first.
This is the net inventory of original files normalized from CRLF to LF: 16 files, 10,402 CRLF pairs removed.
Every listed file changed only by removing CR from CRLF pairs in this phase. No CSV or generated product data changed.
Exact before/after SHA-256 values and byte counts are in [eol-accounting.json](eol-accounting.json).
Later Markdown emphasis/list formatting in the arcade document is separate from this byte-only accounting.

README, LICENSE and the existing VS Code settings were already LF and are excluded from net EOL normalization.
The existing VS Code settings remain byte-identical to the initial working tree; LICENSE is byte-identical to LICENSE.txt.
An intermediate README serialization round-trip returned to LF and contributes no net EOL change.
New template/report text follows the adopted attributes; formatter corrections to those new files are not changes to original data.

| Path | Before bytes | After bytes | CRLF pairs removed |
| --- | ---: | ---: | ---: |
| `ARCADE_EMULATOR_BACKGROUND_INFO.md` | 7924 | 7841 | 83 |
| `Convert-ArcadeBelgiumNplayersIniToCsv.ps1` | 41609 | 40929 | 680 |
| `Convert-ArcadeManagerClassicsListToCsv.ps1` | 4177 | 4099 | 78 |
| `Convert-FBNeoArcadeDATToCSV.ps1` | 25084 | 24595 | 489 |
| `Convert-MAME2003DATToCSV.ps1` | 36738 | 36024 | 714 |
| `Convert-MAME2003PlusDATToCSV.ps1` | 37111 | 36397 | 714 |
| `Convert-MAME2010DATToCSV.ps1` | 39997 | 39216 | 781 |
| `Convert-MAMEDATToCsv.ps1` | 45811 | 44935 | 876 |
| `Convert-ProgettoEmmaCatverIniToCsv.ps1` | 44719 | 43995 | 724 |
| `Convert-ProgettoSnapsBestGamesIniToCsv.ps1` | 43250 | 42541 | 709 |
| `Convert-ProgettoSnapsCategoryIniFilesToCsv.ps1` | 84773 | 83388 | 1385 |
| `Convert-ProgettoSnapsCatverIniToCsv.ps1` | 44341 | 43623 | 718 |
| `Convert-ProgettoSnapsLanguagesIniToCsv.ps1` | 40366 | 39707 | 659 |
| `Convert-ProgettoSnapsRenameSetIniToCsv.ps1` | 35851 | 35215 | 636 |
| `Convert-ProgettoSnapsSeriesIniToCsv.ps1` | 44339 | 43604 | 735 |
| `Create-ConsolidatedAllKillerNoFillerGameList.ps1` | 25066 | 24645 | 421 |
