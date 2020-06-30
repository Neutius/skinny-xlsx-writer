# Software Component: Skinny .xlsx Writer

## Description
The Skinny .xlsx Writer allows you to create a new .xlsx file with a single sheet, add a column header row to the top of this sheet, 
and write rows of data to the sheet. You can add more sheets to the .xlsx and add more rows to the current sheet.

The Skinny .xlsx Writer can only move forward, and has no ability to manipulate sheets other than the newest one. The fanciest 
option is a top row with bold font and freeze pane applied.

When first called, the Skinny .xlsx Writer immediately writes an empty version of the .xlsx file to the target File location 
provided, ensuring any problems writing to disk become apparent as quickly as possible.

This module is meant to be as light-weight as possible. It contains a single class with hardly any methods, and uses a minimum of 
dependencies. It's basically a thin wrapper with some convenience methods and an excessive unit test suite.

The original use case is domain-driven data processing and data analysis, where a medium sized xml-file (max 400MB) is processed.
Parts of the original data, along with derived data and data analysis, is written to an .xlsx file for further analysis or processing.

## Design Goals
This module is designed with two main goals:
- Be as light-weight as possible.
	- There are many powerful libraries for working with .xlsx files, which offer great functionality.
	- For simpler use cases, most of these libraries are simply too bloated.
- Work without a template file.
	- A template file forces you to maintain the code and the template in tandem is error prone, bad practice, and simply annoying.
	- A spreadsheet is just a data container that is slightly more human readable than a database or an xml file.

## API Design
1. Create a new file.
	- Required parameters: target folder location, file name, name of the first sheet.
	    - If the sheet name is null or blank, a valid sheet name will be provided.
	    - If the file name is null or blank, a valid file name will be provided.
	    - If there's any problem with the target folder location, the resulting IOException will not be caught.
	- No need to add the .xlsx extension yourself, it will be added automatically and without double checking.
	    - If the provided file name ends with ".xlsx", the target file name will end with ".xlsx.xlsx".
2. Within the current sheet:
	1. Optional: add column headers to the first row of the sheet.
		- Required parameters: List<String> cell contents.
		    - Any null value will result in a NullPointerException, which will not be caught.
		    - Any blank String is considered an Illegal Argument - a header should contain at least one non-white space character.
		- Has to be added first, and can only be added once per sheet.
		    - Having two column header rows, or have headers below content, is considered an Illegal State.
	2. Add content rows:
		- Required parameters: List<String> cell contents.
		    - This List can be null, can be empty, contain null, or contain blank Strings. 
3. Create a new sheet.
	- Required parameters: sheet name.
	    - If the sheet name is null or blank, or a duplicate of another sheet name, a valid sheet name will be provided.
4. If a new sheet was created: return to 2.
5. Write the .xlsx file to disk.
    - You can do this as often as you like - any earlier versions of the file will be overwritten.
6. There's no need to close the Skinny .xlsx Writer:
    - The in memory Workbook representation has nothing to close.
    - Any OutputStream will be closed immediately after writing to disk.

## Alternate API design - Work in Progress
1. As above, but to everything with one method call.
    - Only functional difference: column header rows have be added to all sheets, or not at all.

