# Comments and Notes Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Comments and Notes


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Comments and Notes**.

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
- create-a-new-workbook-and-add-a-threaded-comment-to-cell-a1-with-author-john.cs
- load-an-existing-workbook-retrieve-all-threaded-comments-from-column-b-and-list-their-authors.cs
- iterate-through-a-threadedcommentcollection-to-display-each-comments-text-author-and-creation-timestamp.cs
- edit-a-specific-threaded-comment-by-setting-its-text-property-to-a-new-string-value.cs
- remove-a-threaded-comment-from-cell-c3-using-the-remove-method-on-the-comment-object.cs
- set-the-text-direction-of-a-comments-shape-to-righttoleft-for-bidirectional-language-support.cs
- set-the-text-direction-of-a-comments-shape-to-toptobottom-for-vertical-annotation-layout.cs
- change-the-font-color-of-a-comment-by-assigning-a-red-value-to-shapetextbodyfontcolor.cs
- update-the-font-color-of-all-comments-authored-by-alice-to-green-using-shapetextbodyfontcolor.cs
- apply-a-solid-blue-background-to-a-comment-using-shapefillforecolor-with-the-appropriate-color-code.cs
- create-a-workbook-add-threaded-comments-to-multiple-cells-and-save-the-file-in-xlsx-format.cs
- load-a-workbook-modify-comment-font-colors-based-on-author-and-save-changes-to-a-new-file.cs
- batch-process-a-folder-of-workbooks-adding-a-standard-disclaimer-comment-to-each-worksheets-top-left-cell.cs
- copy-a-threaded-comment-from-cell-e5-to-cell-f6-while-preserving-its-author-and-text.cs
- read-all-threaded-comments-from-a-worksheet-and-count-the-number-of-comments-per-author.cs
- remove-all-comments-older-than-thirty-days-from-a-workbook-based-on-their-createdtime-values.cs
- scan-a-workbook-for-empty-comments-and-remove-them-to-clean-metadata.cs
- compare-two-workbooks-by-extracting-their-threaded-comments-and-identifying-differences-in-author-attribution.cs
- read-the-author-of-each-threaded-comment-in-a-worksheet-and-output-the-list.cs
- change-the-font-color-of-comments-in-column-g-to-blue-using-shapetextbodyfontcolor.cs
- add-a-threaded-comment-with-multi-line-text-to-cell-h2-and-preserve-line-breaks.cs
- retrieve-and-display-the-total-number-of-threaded-comments-present-in-a-workbook.cs
- update-the-text-direction-of-all-comments-in-a-worksheet-to-lefttoright-for-standard-layout.cs
- replace-the-background-picture-of-a-comment-with-a-semi-transparent-overlay-image.cs
