
This repository contains several Excel VBA macros for parsing, transforming, joining, and splitting text data.

The tools are useful when you need to:

* parse values from Column A
* extract a fixed part of each string
* replace text using regular expressions
* join many rows into one long text
* split long text into safe-sized Excel cells
* convert values to uppercase and lowercase variants

## Included macros

### 1) `VBA_replace`

Main parsing macro.

**Macro name:** `ColumnA_ToRowString_SplitToColumnB`

What it does:

* reads data from **Column A**
* checks for duplicates in Column A
* asks which parser method to use:

  * `1 = Fixed MID extract`
  * `2 = Regex replace`
* asks for separator
* asks for max cell length
* joins parsed values into one long string
* splits the result into chunks
* writes output to **Column C**

**Important:** despite the macro name mentioning “ColumnB”, the current code writes to **Column C**.

#### Method 1: Fixed MID extract

Use this when you want to extract a fixed part of each text.

Example:

* source value: `INV-2026-000123`
* start position: `1`
* width: `8`
* result: `INV-202`

#### Method 2: Regex replace

Use this when you want to search or remove part of each string with a regex pattern.

Example:

* source value: `INV-2026-000123`
* pattern: `INV-`
* replace with: empty string
* result: `2026-000123`

### 2) `VBA_join`

**Macro name:** `ColumnA_TransposeAndSplitToColumnC_WithOptionalSides`

What it does:

* reads non-empty values from **Column A**
* asks for separator between values
* asks for optional left wrapper
* asks for optional right wrapper
* asks for max output cell length
* joins all values together
* splits the long result into chunks
* writes output into **Column C**

Example:

* Column A:

  * apple
  * banana
  * orange
* separator: `,`
* left side: `[`
* right side: `]`

Result:
`[apple],[banana],[orange]`

### 3) `VBA_stacktherows`

**Macro name:** `ColumnA_TransposeAndSplitToColumnC`

What it does:

* reads non-empty values from **Column A**
* asks for separator
* asks for max cell length
* joins all values into one string
* splits the result into multiple cells if needed
* writes output into **Column C**

This is the simpler version of `VBA_join` without optional left/right wrappers.

### 4) `VBA_textToColumn_transpouse`

**Macro name:** `SplitColumnA_ToColumnC`

What it does:

* asks for a separator
* reads each non-empty cell in **Column A**
* splits each cell using the separator
* writes each split item vertically to **Column C**

Example:

* A1 = `dog,cat,bird`
* separator = `,`

Output in Column C:

* C1 = `dog`
* C2 = `cat`
* C3 = `bird`

### 5) `VBA_caps_revert`

**Macro name:** `RowToColumn_UpperLower`

What it does:

* reads non-empty cells from **Row 1**
* for each value, writes:

  * uppercase version
  * lowercase version
* outputs the results vertically in **Column A**

Example:

* Row 1 contains: `Apple`, `Banana`

Output in Column A:

* `APPLE`
* `apple`
* `BANANA`
* `banana`

## How to install in Excel

1. Open Excel.
2. Press `Alt + F11` to open the VBA editor.
3. In the VBA editor, go to **Insert > Module**.
4. Copy the code from one of the repository files.
5. Paste it into the module.
6. Save the workbook as **Excel Macro-Enabled Workbook (`.xlsm`)**.

## How to run a macro

1. Open the workbook with your data.
2. Make sure macros are enabled.
3. Press `Alt + F8`.
4. Select the macro you want to run.
5. Click **Run**.

## Expected input layout

Most macros in this repository use **Column A** as the source.

Typical input example:

| A            |
| ------------ |
| INV-2026-001 |
| INV-2026-002 |
| INV-2026-003 |

Depending on the macro, output is usually written to **Column C**.

* Regex logic uses `VBScript.RegExp`, so regex syntax follows the VBScript regular expression engine. ([GitHub][1])

## Quick start example

If you want to parse values from Column A and combine them into safe-length cells:

1. Open `VBA_replace`
2. Run `ColumnA_ToRowString_SplitToColumnB`
3. Choose:

   * `1` for fixed extract, or
   * `2` for regex replace
4. Enter separator, for example `,`
5. Enter max cell length, for example `256`
6. Check the result in **Column C**

## License

This repository is published under the **MIT License**.
