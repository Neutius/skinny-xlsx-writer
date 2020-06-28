# Software Component: Skinny xlsx Writer

## Description
The Skinny xlsx Writer allows you to create a new .xlsx file with a single sheet and write rows of data to this sheet.

The two optional possibilities are adding column headers to a sheet, and adding a title or description

This module is meant to be as light-weight as possible.

The original use case is domain-driven data processing and data analysis, where a medium sized xml-file (max 400MB) is processed.
Parts of the original data, along with derived data and data analysis, is written to an .xlsx file for further analysis or processing.

## Design Goals
This module is designed with two main goals:
- Be as light-weight as possible.
	- There are many powerful libraries for working with .xlsx files, which offer great functionality.
	- For simpler use cases, most of these libraries are simply too bloated.
- Work without a template file.
	- Being forced to maintain the code and the template in tandem is error prone, bad practice, and simply annoying.


## API Design

1. Create a new file.
	- Required parameters: target folder location, file name, name of the first sheet.
2. Within the current sheet:
	1. Optional: add a title or description to the first row.
		- Required parameters: title or description.
		- Optional parameters: number of cells to merge (default: 5 cells).
	2. Optional: add column headers to the first or second row of the sheet.
		- Required parameters: String[] cell contents.
	3. Add content rows:
		- Required parameters: String[] cell contents.
3. Create a new sheet.
	- Required parameters: sheet name.
4. If a new sheet was created: return to 2.
5. Write the .xlsx file to disk.


## Implementation Design (note to self, don't push this part to git)

- Use File API or use Path API?

- Sheet:
	- standard no column header (boolean: false). If added: boolean true + String[]
	- same for title/description
	- fields: boolean withTitle, String title, boolean withHeaders, String[] 



