# Master Report Center v2.0.2 — Complete User Guide

**Publisher:** Golden Fir Solutions  
**Platform:** Windows 10 / Windows 11  
**Version:** 2.0.2  
**Product:** Master Report Center All In One

Master Report Center is a Windows desktop application for Excel reporting and data workflows. It includes four core modules:

- Split Reports
- Consolidate Reports
- Synchronize Reports
- Lookup Data

This guide explains installation, licensing, normal operation, module behaviour, data-safety rules, output files, and common troubleshooting.

---

## 1. System Requirements

### Required

- Windows 10 or Windows 11
- Microsoft Excel Desktop
- Internet access for purchase, activation, recovery, refund requests, and licence-status synchronization

Microsoft Excel Desktop is required because several workflows use Excel COM automation to preserve workbook structure, formatting, worksheet protection settings, and other Excel-native behaviour.

### Performance

Processing speed depends on the computer, Excel version, workbook size, formulas, formatting complexity, file location, antivirus scanning, available memory, and storage speed. Large workbooks may run noticeably faster on a higher-performance office computer than on a home PC.

For best results:

- Keep sufficient free disk space available.
- Avoid editing the same workbook while MRC is processing it unless the module specifically requires the currently open Excel view.
- Prefer local or fast network storage when possible.
- Allow Excel background processing to finish before starting another large task.

---

## 2. Installation and Startup

### Standard Installer

The recommended package is:

`Master_Report_Center_v2.0.2_Setup.exe`

The installer creates the application installation, Start Menu integration, and normal Windows uninstall support.

### Portable Build

The portable package is:

`Master_Report_Center_v2.0.2_Portable.zip`

Extract the entire ZIP before launching `Master Report Center.exe`. Do not run the executable directly from inside the ZIP archive.

### Local Application Data

Runtime state is stored outside the installation folder under:

`%LOCALAPPDATA%\GoldenFir\MasterReportCenter\`

Reinstalling the application does not automatically remove local trial or licence state.

---

## 3. Trial, Purchase, Activation and Recovery

### Trial

Master Report Center includes a **12-day trial**. The License page shows the remaining trial period and current licence state.

When the trial expires, report modules are locked until a valid licence is purchased, activated, or recovered.

### Licence Options

**Personal License — AUD 49 one-time**

- Individual professional use
- Up to 4 device migrations or rebindings in a rolling 12-month period
- One-time purchase, no subscription

**Company License — AUD 149 one-time**

- Business-owned use
- No annual device recovery / rebinding count limit
- Migration and recovery events are still recorded
- One-time purchase, no subscription

### Activation

Use **License → Activate License** when activating an existing paid licence on the current installation.

Keep the purchase email and Recovery Code in a safe place. They are important for future account recovery and permitted device transfer.

### Recovery

Use **Recover Existing Account** when restoring an existing licence.

A same-device restoration does not count as a new migration. A genuinely new installation identity may count as a migration under the Personal License rules.

### Refunded or Revoked Licences

If a refund is approved and confirmed by the payment provider, the licence is revoked. Existing reports and files remain on the computer, but MRC report modules become locked.

---

## 4. Main Application Areas

### Home

The Home page provides access to the four report modules and general status information.

Typical Home information includes:

- module launch cards
- Recent Projects / recent activity
- Tips & Guidance
- System Status, including trial or licence state

### License

Use the License page to:

- view trial or licence status
- purchase a licence
- activate an existing licence
- recover an existing account
- copy purchase email / Recovery Code when available
- view Personal licence migration allowance
- submit a refund request when eligible

### Settings

Settings includes **Configuration backup & transfer** and diagnostics.

**Export Configuration** creates a JSON backup of supported settings.

**Import Configuration** loads a previously exported configuration file. Review the source of imported configuration files before use because importing can change application behaviour.

**View Log Folder** opens the diagnostic log location for troubleshooting.

### About

The About page shows the application version, publisher, and the four supported workflows.

---

# 5. Split Reports — Complete Guide

## 5.1 Purpose

Split Reports divides one large Excel report into multiple smaller workbooks based on the unique values in a selected column.

Typical uses include splitting by:

- Supplier
- Department
- Country
- Store
- Category
- Team
- Status

Supported source formats include `.xlsx`, `.xlsm`, and `.xls`.

## 5.2 Standard Workflow

1. Select the **Source Excel File**.
2. Select the **Output Folder**.
3. Select the worksheet.
4. Confirm the Header Row.
5. Select the **Split Column**.
6. Choose the required data mode.
7. Click **Analyze Report**.
8. Review row count, output count, blank split values, and filename preview.
9. Configure optional filename suffix and date suffix.
10. Click **Start Splitting**.
11. Monitor the progress bar and Activity Log.
12. When complete, MRC opens the output folder.

## 5.3 Header Row

The default is **Auto Detect**. MRC can also be manually set to a header row from Row 1 to Row 50.

If the workbook contains title rows, notes, or other information above the actual report headers, confirm that the detected row is correct before processing.

## 5.4 Full Workbook Data vs Current Excel Workbook View

### Full Workbook Data

Reads all data rows from the selected worksheet. Saved filters and hidden-row state are not treated as the active processing set.

Use this mode when the entire worksheet should be split.

### Current Excel Workbook View

Processes only rows currently visible in the open Excel workbook after filtering or row hiding.

Use this mode when you have already filtered the report in Microsoft Excel and only want the visible records included.

**Important:** Current View must be analyzed first. MRC creates an Excel-native temporary view copy and verifies the row and output-group counts before splitting.

If the file, worksheet, split column, header row, or view mode changes after analysis, analyze again before starting the split.

## 5.5 Protected Worksheets

A protected worksheet cannot be processed in Full Workbook mode.

For a protected worksheet:

1. Open the workbook in Microsoft Excel.
2. Apply the required filters or row hiding.
3. Enable **Use Currently Open Excel Workbook View**.
4. Click **Analyze Report**.
5. Enter the worksheet password if requested.

The password is used only during the current MRC session and is not saved or written to the Activity Log.

Incorrect passwords do not modify the source workbook or create split files.

## 5.6 Options

### Exclude Hidden Columns

Enabled by default.

Hidden columns are normally excluded from output. However, if MRC detects dropdown Data Validation that depends on hidden supporting columns, required supporting columns may be retained automatically.

### Include Other Worksheets

Enabled by default.

When enabled, each split output keeps the other worksheets in the source workbook.

If disabled, MRC keeps only the selected split worksheet. This may affect:

- cross-sheet formulas
- Named Ranges
- charts
- dropdown / Data Validation sources

MRC displays a warning before continuing when this option is turned off.

## 5.7 Analyze Report

Analysis reports information such as:

- Source mode
- Worksheet
- Header row
- Split column
- Rows / visible rows
- Columns
- Number of output files to create
- Blank split values
- Worksheet protection status
- Filename example

Use the analysis result to confirm the task before creating files.

## 5.8 Output Naming

Output filenames use the split value and can also include:

- a custom filename suffix
- the current date as the final suffix

Example:

`Supplier A - Testing Outstanding - 10-10-2026.xlsx`

The date suffix is enabled by default.

---

# 6. Consolidate Reports — Complete Guide

## 6.1 Purpose

Consolidate Reports combines multiple related Excel reports from one folder into a single consolidated workbook.

Supported source formats:

- `.xlsx`
- `.xlsm`
- `.xlsb`
- `.xls`

Temporary Excel files beginning with `~$` and previously generated `Combined_Report_...` files are excluded automatically.

## 6.2 Standard Workflow

1. Select the **Source Reports Folder**.
2. Select the **Output Folder**.
3. Optionally enter a worksheet name.
4. Review Data Handling options.
5. Click **Analyze Reports**.
6. Confirm the Excel files found.
7. Click **Start Consolidation**.
8. Microsoft Excel runs in the background.
9. Review the Activity Log and final Consolidate Log.

## 6.3 Worksheet Selection

If the worksheet field is blank, the first worksheet is used.

If a worksheet name is entered, that same worksheet name must exist in each source workbook intended for processing.

## 6.4 Header Rules

All valid source files must have the same **normalized header set**, but the column order may differ.

For example, these can still be consolidated:

`Style | Supplier | Country | Qty`

and

`Country | Qty | Style | Supplier`

MRC maps columns by header name.

A file may be skipped when:

- a required header is missing
- an extra header changes the expected header set
- a header is blank
- duplicate header names exist
- the selected worksheet is missing
- the workbook cannot be opened

Skipped files and reasons are recorded in the log.

## 6.5 Output Template

The first valid source workbook is copied and used as the output format template.

The existing data area is replaced with the consolidated result while Excel-native copy/paste is used to preserve formatting and workbook structure as faithfully as possible.

## 6.6 Data Handling Options

### Ignore completely blank rows

Enabled by default. Fully blank rows are not written to the combined report.

### Remove exact duplicate rows

Enabled by default. Rows with exactly equivalent values are kept only once.

### Add source filename column

Disabled by default. When enabled, MRC adds a source filename field so each consolidated record can be traced back to its original workbook.

## 6.7 Protected Template Worksheet

If the selected output-template worksheet is protected, MRC requests the worksheet password.

The password is not saved or logged.

Password validation is performed using a temporary copy. During the actual task, MRC temporarily unlocks only the copied output workbook and then restores the worksheet protection settings before saving the result.

The original source workbooks are not modified.

## 6.8 Output and Log

Typical output:

`Combined_Report_Merged_2026-10-10_093000.xlsx`

Typical log:

`Consolidate_Log_2026-10-10_093000.txt`

The log records items such as:

- source and output folders
- worksheet
- selected Data Handling options
- protection status
- source files found / processed
- skipped file details
- source rows read
- rows written
- duplicates removed
- blank rows skipped
- output file
- duration
- final status

---

# 7. Synchronize Reports — Complete Guide

## 7.1 Purpose

Synchronize Reports uses one or more Source reports to update a Destination / Master workbook.

It is designed for **match-and-update**, not simple file concatenation.

The workflow is:

`Match existing row → Update selected fields → Optionally append new rows`

Supported formats include `.xlsx`, `.xlsm`, `.xlsb`, and `.xls`.

## 7.2 Destination and Source

**Destination** is the Master workbook that needs updated data.

**Source** is a folder containing one or more workbooks that provide the newer values.

The original Destination file is never overwritten. MRC creates an updated copy.

Files already containing `_Updated_` are excluded from Source processing, as are Excel temporary files and the Destination workbook itself.

## 7.3 Standard Workflow

1. Select the Destination workbook.
2. Select the Destination worksheet.
3. Select the Source reports folder.
4. Select the Source worksheet.
5. Click **Analyze Headers**.
6. Choose one or more **Match Fields**.
7. Choose one or more **Write Fields**.
8. Review synchronization options.
9. Click **Start Synchronization**.
10. Review the generated Updated workbook and Synchronize Log.

## 7.4 Worksheet Selection

Destination and Source worksheets are selected separately.

Visible, Hidden, and VeryHidden worksheets can appear in the selector. The Source worksheet list is sampled from the first valid Source workbook; the selected worksheet name must exist in each Source file intended for processing.

## 7.5 Analyze Headers

Header analysis is required before synchronization can start.

It identifies:

- Destination headers
- Source headers
- Common headers available for matching / writing
- Destination worksheet
- Source worksheet
- Source sample file

Blank or duplicate headers are invalid.

## 7.6 Match Fields

Match Fields identify which Source row corresponds to which Destination row.

Use stable identifiers such as:

- Style Code
- SKU
- Employee ID
- Order ID
- Certificate Number

Multiple Match Fields can be combined, for example:

`Style Code + Country`

Text key matching is normalized for surrounding spaces and case.

Fields expected to change should normally **not** be used as Match Fields.

## 7.7 Write Fields

Write Fields are the values Source is allowed to write into a matched Destination row.

Example:

Match Field:

`Style Code`

Write Fields:

`Status, Expiry Date, Supplier, Comment`

The same field cannot be selected as both a Match Field and a Write Field.

Updates are values-only so Destination formatting remains intact.

## 7.8 Duplicate Destination Keys

If the Destination contains multiple rows with the same Match key, the **first Destination row wins**.

The duplicate-key count is reported in the log and should be reviewed if uniqueness is expected.

## 7.9 Synchronization Options

### Add unmatched source rows

Enabled by default.

If a Source key does not exist in Destination, MRC appends a new row. For that new row, all Source fields with matching Destination headers can be copied.

If disabled, unmatched Source rows are not written and are counted as unmatched.

### Ignore blank source values when updating matched rows

Enabled by default.

A blank Source value does not overwrite an existing Destination value.

Example:

Destination Status: `Approved`  
Source Status: blank  
Result: `Approved`

## 7.10 Protected Destination Worksheet

If the Destination worksheet is protected, MRC requests the password.

The password is not saved or logged. Validation uses a temporary copy. During synchronization, only the copied output workbook is unlocked, updated, and then re-protected.

The original Destination workbook remains unchanged.

## 7.11 Output and Log

Typical output:

`MasterRegister_Updated_2026-10-10_093000.xlsx`

Typical log:

`Synchronize_Log_2026-10-10_093000.txt`

The log records:

- Destination file
- Source folder
- Destination / Source worksheet
- Match Fields
- Write Fields
- append-unmatched option
- ignore-blank option
- protection status
- files found / processed / skipped
- source rows read
- matched rows
- updated rows
- matched rows with no write
- appended rows
- unmatched rows
- duplicate Destination keys
- output file
- duration and status

---

# 8. Lookup Data — Complete Guide

## 8.1 Purpose

Lookup Data enriches a smaller Target workbook using selected fields from a larger Reference Database.

It is designed for:

`Large Reference Database → Smaller Target workbook`

The module reads only the required Match and Return fields and writes values into the Target copy while preserving Target formatting.

Supported formats:

- `.xlsx`
- `.xlsm`
- `.xlsb`
- `.xls`

## 8.2 Reference vs Target

**Reference Database** contains the larger information set and the values to retrieve.

**Target** is the workbook that needs enrichment.

The original Target workbook is never overwritten. MRC creates a new Lookup output copy.

Lookup does not append unrelated Reference records to Target.

## 8.3 Standard Workflow

1. Select the Reference Database workbook.
2. Optionally enter the Reference worksheet name; blank uses the first worksheet.
3. Select the Target workbook.
4. Optionally enter the Target worksheet name; blank uses the first worksheet.
5. Click **Analyze Headers**.
6. Choose one or more Match Fields.
7. Choose one or more Return Fields.
8. Review Safe Write Behaviour.
9. Click **Start Lookup**.
10. Review the Lookup output workbook and Lookup Log.

## 8.4 Analyze Headers

Header analysis compares the Reference and Target files.

**Match Fields** must exist in both files.

**Return Fields** come from the Reference Database and represent the values to retrieve.

Blank or duplicate header names are invalid.

## 8.5 Match Fields

Match Fields identify which Reference record belongs to each Target row.

Example:

`Style Code`

or a composite key such as:

`Style Code + Country`

Text matching is normalized for case and surrounding spaces.

A Match Field cannot also be selected as a Return Field.

## 8.6 Return Fields

Return Fields are values copied from the matched Reference record into the Target output.

Examples:

- Supplier Name
- Department
- Category
- Registration Number
- Status
- Expiry Date

If a selected Return Field does not already exist in the Target, MRC can add the missing Target field to the output workbook.

## 8.7 Safe Write Behaviour

### Only fill blank Target cells

Enabled by default.

Existing non-blank Target values are preserved. This is the safest mode when Target already contains manually reviewed information.

If disabled, matched Reference values may replace existing Target values in selected Return Fields.

### Ignore blank Reference values

Enabled by default.

Blank Reference values do not clear Target values.

## 8.8 Duplicate Reference Keys

If the Reference contains duplicate Match keys, the duplicate count is recorded in the Lookup Log. Review duplicate keys if the Reference Database is expected to contain unique identifiers.

## 8.9 Output and Log

Typical output:

`TargetReport_Lookup_2026-10-10_093000.xlsx`

Typical log:

`Lookup_Log_2026-10-10_093000.txt`

The log records:

- Reference Database
- Target workbook
- Reference / Target worksheet
- Match Fields
- Return Fields
- safe-write options
- Reference rows read
- Target rows read
- matched Target rows
- unmatched Target rows
- cells written
- existing Target cells preserved
- blank Reference values skipped
- duplicate Reference keys
- missing Target fields added
- output file
- duration and status

---

# 9. Data Safety and Recommended Practice

Master Report Center is designed to reduce accidental source-file changes, but users should still maintain normal business backups.

Recommended practice:

1. Keep an original copy of important workbooks.
2. Review Header Analysis / Report Analysis before running a large task.
3. Use stable identifiers as Match Fields.
4. Read Activity Log warnings instead of assuming every Source file was processed.
5. Review persistent module logs after large or business-critical jobs.
6. Do not delete output logs until the result has been checked.
7. For protected workbooks, use the correct worksheet password and verify protection on the output.

Important module-specific safety behaviour:

- Split Reports creates separate output workbooks.
- Consolidate Reports builds a new combined workbook.
- Synchronize Reports never overwrites the original Destination.
- Lookup Data never overwrites the original Target.

---

# 10. Common Warnings and Troubleshooting

## Microsoft Excel is required

If MRC reports that Microsoft Excel or pywin32 is unavailable, confirm that Microsoft Excel Desktop is installed and working normally on Windows.

## Excel is busy

MRC automatically retries certain temporary Excel COM busy conditions. If the computer is heavily loaded, allow Excel to finish background work and retry.

## Processing is slower than expected

Possible causes include:

- large numbers of rows or files
- complex workbook formatting
- formulas or external links
- protected worksheets
- slower CPU / storage
- limited memory
- network-drive latency
- antivirus scanning
- other Excel processes running at the same time

A slower runtime does not by itself indicate incorrect processing. Use the Activity Log and output log to confirm completion and row/file counts.

## Header mismatch

Check that required source files use the expected fields. Column order can often differ, but missing, extra, blank, or duplicate headers may cause a file to be skipped or analysis to fail.

## Worksheet not found

Confirm the selected worksheet name exists in the workbook. For Consolidate and Synchronize, the same selected Source worksheet must exist in each relevant Source workbook.

## Protected worksheet password rejected

Confirm the worksheet password in Microsoft Excel. MRC does not save worksheet passwords.

## Output appears incomplete

Review the persistent module log for:

- files skipped
- unmatched rows
- duplicate keys
- blank values ignored
- rows removed as duplicates
- visible-row-only processing

---

# 11. Understanding the Four Modules

| Module | Main Purpose | Source | Result |
|---|---|---|---|
| Split Reports | Divide one report by a selected field | One workbook | Many workbooks |
| Consolidate Reports | Merge related reports | Folder of workbooks | One combined workbook |
| Synchronize Reports | Match and update a master | Destination + Source folder | Updated copy of Destination |
| Lookup Data | Enrich a smaller workbook from a larger database | Reference + Target | Enriched copy of Target |

Choose the module based on the business task rather than the number of files alone.

---

# 12. Support Information to Keep When Reporting a Problem

When requesting support, keep the following information available:

- Master Report Center version
- Windows version
- Microsoft Excel version
- module being used
- approximate row count / file count
- Activity Log message
- persistent module log
- whether the workbook is protected
- whether files are local, network-based, OneDrive/SharePoint-synced, or on removable storage

Do **not** send passwords or private licence Recovery Codes in ordinary diagnostic messages.

---

## Version Note

This guide documents **Master Report Center v2.0.2** and is intended to match the public v2.0.2 release behaviour. Future versions may change controls, processing rules, licensing behaviour, or supported workflows.

**Golden Fir Solutions**  
*Automate • Simplify • Empower*
