# Workbook Merger Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Workbook Merger


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Workbook Merger**.

Example:

create-a-workbook.cs


## Required Namespaces

Most examples will require:

using Aspose.Cells;


## Common Pattern

Typical Aspose.Cells workflow:

Workbook workbook = new Workbook();

Worksheet sheet = workbook.Worksheets[0];

Cells cells = sheet.Cells;


## Output

Examples may generate:

- XLSX files
- PDF files
- CSV files
- Images

Output files are written to the working directory.
- add-the-library-reference-to-a-net-project-via-nuget-before-implementing-merge-logic.cs
- create-a-console-application-project-in-visual-studio-to-host-the-workbook-merging-code.cs
- load-source-xls-workbooks-using-new-workbookfilepath-for-each-file-to-be-merged.cs
- load-large-xls-files-with-cellshelpermergefiles-by-providing-an-array-of-file-paths-and-output-path.cs
- use-workbookcombine-to-merge-two-or-more-workbooks-when-file-sizes-are-moderate.cs
- copy-specific-worksheets-from-source-workbooks-into-the-target-workbook-using-worksheetcopy-method.cs
- preserve-charts-and-images-during-merge-by-employing-default-workbookcombine-behavior-without-additional-options.cs
- ensure-formulas-remain-intact-by-keeping-calculation-mode-set-to-automatic-before-and-after-merging.cs
- maintain-original-cell-formatting-by-not-altering-style-settings-during-the-combine-operation.cs
- verify-that-the-merged-workbook-contains-the-expected-number-of-worksheets-after-combination.cs
- check-that-all-charts-from-source-workbooks-appear-correctly-in-the-combined-workbook.cs
- confirm-that-embedded-images-are-retained-in-the-merged-workbook-after-using-workbookcombine.cs
- save-the-merged-workbook-to-a-specified-output-path-using-workbooksave-method.cs
- export-the-combined-workbook-to-pdf-format-to-verify-visual-fidelity-of-charts-and-images.cs
- export-each-worksheet-of-the-merged-workbook-to-csv-files-for-data-extraction-validation.cs
- generate-an-html-representation-of-the-merged-workbook-to-inspect-content-in-a-web-browser.cs
- save-the-merged-workbook-into-a-memory-stream-for-immediate-transmission-via-a-web-api.cs
- attach-the-merged-workbook-file-to-an-email-using-systemnetmail-after-successful-combination.cs
- encrypt-the-merged-workbook-with-a-password-before-saving-to-protect-confidential-data.cs
