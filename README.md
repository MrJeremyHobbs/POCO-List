![Alt text](https://github.com/MrJeremyHobbs/POCO-List/blob/master/images/poco_list_1_5_screenshot.JPG?raw=true "Screenshot") 
# POCO-List
POCO-List takes a spreadsheet from Alma and generates a nightly inventory for Cal Poly Kennedy Library's POCO lab.

Check releases tab for latest distribution.

## Troubleshooting
Every once and awhile Alma changes the columns for the input spreadsheet. PocoList looks for specific data in specific columns, so these changes will break  the program.

In the config.ini file, there are three entries for three columns in the spreadsheet. Each entry is the number of the column in Excel. When Alma makes changes to the spreadsheet, look at the input file and find the correct column numbers and change the config file to match.

By default, Excel uses letter notation for their columns. To see the number format:
-Open input spreadsheet
-Click File->Options
-Choose "Formulas" from left-hand menu
-Under "Working with formulas", click "R1C1 reference"

## Editing the Config.ini File
Find the new column number and enter into the config.ini file and re-run.

```[spreadsheet]
path=c:\Temp

[columns]
;must be in number format, not letter
description_col=34
title_col=2
permanent_location_col=19

[misc]
delete_when_done=active