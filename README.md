# Excel2MySQL

**Convert an Excel worksheet into a MySQL `.sql` script—locally, without uploading your workbook.**

Excel2MySQL is a Windows desktop application for previewing an `.xlsx` or `.xlsm` worksheet and exporting its data as MySQL `INSERT` statements. It does not connect to a database or execute SQL: you review and run the generated script with your preferred MySQL client.

> **Privacy:** Workbook processing and SQL generation happen on your computer. The application does not upload files or require an internet connection to convert them. Internet access is needed to install Python packages or download a release.

## Features

- Open an Excel workbook with the file picker or drag and drop.
- Select a worksheet and preview its headers and a configurable number of rows.
- Preview the generated SQL and update it as options change.
- Choose a MySQL table name or generate a suggested name from the workbook and worksheet.
- Optionally include `DROP TABLE IF EXISTS` and `CREATE TABLE` statements.
- Choose batched or individual `INSERT` statements and set the batch size.
- Export the complete script as a UTF-8 `.sql` file, including for large worksheets.
- Cancel an export in progress. A cancelled or failed export does not leave a partial output file.
- Use the graphical application without installing Python by downloading the Windows executable from a release.

## Download and run

### Windows executable

1. Open the repository’s [Releases](https://github.com/GayanPradeepJayawardhana/Excel2SQL/releases) page.
2. Download `Excel2MySQL.exe` from the release assets.
3. Run the executable and open a workbook.

The executable is a Windows desktop build. Python is not required to run it. If no executable asset is available on the Releases page, use the source instructions below or build the executable yourself.

### Run from source

Use Python on Windows. The required packages are listed in [`requirements.txt`](requirements.txt).

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python app.py
```

If PowerShell prevents environment activation, install the packages using the virtual environment's interpreter and run the app directly:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
```

## Build the Windows executable

From a Windows command prompt or PowerShell, run:

```bat
build.bat
```

The script installs or updates the dependencies from `requirements.txt` and uses PyInstaller to create a one-file, windowed application at:

```text
dist\Excel2MySQL.exe
```

Building requires Python and an internet connection for package installation (unless all required packages are already available). Running the resulting executable does not require Python.

## How to use

1. Launch Excel2MySQL and open or drop an `.xlsx` or `.xlsm` workbook into the application.
2. Select the worksheet to convert. The first worksheet is selected initially.
3. Review the Excel preview. Adjust **Preview rows** to show between 10 and 500 rows.
4. Enter a MySQL table name, or select **Auto** to suggest one from the workbook and worksheet names.
5. Choose whether to include **DROP TABLE first**, **CREATE TABLE**, and **BATCH INSERT**. Set **Rows per INSERT** if batching is enabled.
6. Review the live SQL preview. It shows only a sample, not the complete output.
7. Select **Save .sql file** and choose where to write the full script. **Convert & Preview** displays generated SQL in the application, up to a 200 KB limit.
8. Review the saved SQL before running it with your MySQL client.

### Important: review destructive SQL

**DROP TABLE first is enabled by default.** If left enabled, the generated script begins with `DROP TABLE IF EXISTS` for the selected table. Running that script can permanently delete an existing table and its data. Disable this option unless you intentionally want to replace the table, and always inspect the SQL before executing it.

The application does not create a database, choose a database/schema, connect to MySQL, or run the generated statements. Select the target database in your MySQL client before executing the script.

## Conversion behavior and limitations

- Supported workbook extensions are `.xlsx` and `.xlsm`. Legacy `.xls` files are not supported.
- One selected worksheet is exported per SQL file. The first row is treated as the column-header row; an empty worksheet or one without usable headers cannot be converted.
- Every column from the first column through the last non-empty header is included. Empty headers between named columns become names such as `column_b`; trailing columns after the last non-empty header are omitted.
- MySQL column-name constraints are applied: duplicate names (case-insensitive) receive suffixes such as `_2`, trailing spaces are removed, and names are limited to 64 characters. Each changed header is reported in a `-- NOTE:` comment.
- When **CREATE TABLE** is enabled, columns are defined as `LONGTEXT` using InnoDB and `utf8mb4`. No primary key or extra ID column is added.
- Blank rows between data are included as rows of `NULL` values. Blank rows at the end of the worksheet are omitted.
- Text is not trimmed. The converter escapes text for SQL string literals and preserves content such as whitespace, line breaks, leading zeroes, and Unicode characters.
- Excel numeric cells are limited by Excel's own numeric precision. A value already rounded by Excel cannot be recovered by the converter.
- Formula cells use the last value saved in the workbook. If the workbook has no cached formula result (for example, it was generated by a script and never recalculated by Excel), the exported formula value may be empty.
- `.xlsm` files are read as workbooks; macros are not run or included in the SQL output.
- Dates and times are emitted as SQL string values. Check that the target MySQL column types and application expectations match the values in your workbook.
- The generated script sets the connection character set to `utf8mb4` and temporarily disables foreign-key checks while inserting rows, then re-enables them.
- Batch size can be set from 1 to 2,000 rows per `INSERT` in the interface. With batching disabled, rows are emitted as individual `INSERT` statements.

## Requirements

- **To run the prebuilt executable:** Windows; Python is not needed.
- **To run or build from source:** Python and the packages in [`requirements.txt`](requirements.txt).
- **To import the generated SQL:** A compatible MySQL server/client and permission to create, drop, or insert into the selected table, as applicable to the options used.

## Troubleshooting

- **The workbook will not open:** Confirm it is a valid `.xlsx` or `.xlsm` file. Excel2MySQL does not support `.xls`.
- **Formula values are empty:** Open the workbook in Excel or another spreadsheet application, recalculate it, save it, and convert it again.
- **The table name is rejected:** Use up to 64 letters, digits, or underscores, and do not start the name with a digit.
- **The output is large:** The live preview is intentionally limited. Use **Save .sql file** to create the full script; conversion is streamed rather than accumulated as one large in-memory SQL string.
- **The SQL fails when imported:** Review the generated script and the selected database, permissions, MySQL version, column constraints, and SQL mode. The optional `DROP TABLE` statement may remove an existing table.
- **The build fails:** Check that Python is installed and available as `python` in the terminal, then rerun `build.bat` and review the reported pip or PyInstaller error.

## Releasing

This repository does not define an application version or automated release workflow. Before publishing a release:

1. Choose a release version and use the same version for the Git tag (for example, `v1.0.0`) and the GitHub release.
2. Build on Windows with `build.bat`.
3. Test `dist\Excel2MySQL.exe` on a Windows machine, including opening a workbook, switching worksheets, changing SQL options, and saving/importing a sample `.sql` file.
4. Create a GitHub release for the version tag and attach `dist\Excel2MySQL.exe` as a release asset.
5. Include concise release notes describing changes and any known limitations. Do not describe the release as tested on platforms or configurations that were not tested.

## Project files

- [`app.py`](app.py) — desktop user interface.
- [`converter.py`](converter.py) — workbook reading and SQL generation.
- [`requirements.txt`](requirements.txt) — Python dependencies.
- [`build.bat`](build.bat) — Windows executable build script.
- [`Excel2MySQL.spec`](Excel2MySQL.spec) — PyInstaller configuration.

## License

No license file is currently included in this repository, so the project does not specify a license for reuse or redistribution.

---

Documentation last updated: **2026-11-09**
