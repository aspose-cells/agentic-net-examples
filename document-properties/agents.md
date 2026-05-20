# Document Properties Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Document Properties


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Document Properties**.

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
- load-a-workbook-and-retrieve-the-title-builtin-property-for-verification.cs
- open-an-excel-file-and-read-the-author-builtin-property-to-identify-the-creator.cs
- instantiate-a-workbook-and-obtain-the-documentversion-builtin-property-to-check-version-information.cs
- load-a-spreadsheet-and-read-the-language-builtin-property-to-determine-locale-settings.cs
- open-a-workbook-and-query-the-scalecrop-builtin-property-to-view-image-scaling-flag.cs
- load-an-excel-document-and-inspect-the-linksuptodate-builtin-property-for-hyperlink-status.cs
- create-a-workbook-and-set-the-title-builtin-property-to-a-descriptive-project-name.cs
- open-a-file-and-update-the-author-builtin-property-with-the-correct-contributor-identifier.cs
- load-a-workbook-and-assign-the-documentversion-builtin-property-the-value-20.cs
- open-a-spreadsheet-and-set-the-language-builtin-property-to-fr-fr-for-localization.cs
- instantiate-a-workbook-and-enable-the-scalecrop-builtin-property-to-preserve-image-proportions.cs
- load-a-workbook-and-disable-the-linksuptodate-builtin-property-to-prevent-link-checks.cs
- open-a-workbook-and-add-a-custom-property-processeddate-with-the-current-datetime-value.cs
- load-an-excel-file-and-create-a-custom-property-projectid-with-an-integer-identifier.cs
- open-a-workbook-and-add-a-custom-boolean-property-isreviewed-set-to-true.cs
- load-a-workbook-locate-the-custom-property-projectid-and-update-its-integer-value.cs
- open-a-spreadsheet-and-remove-the-custom-property-isreviewed-to-clean-obsolete-metadata.cs
- load-a-workbook-and-check-whether-a-custom-property-clientname-exists-before-adding.cs
- open-a-file-and-iterate-through-all-builtin-properties-logging-each-name-and-value.cs
- load-a-workbook-and-enumerate-custom-properties-exporting-their-names-types-and-values-to-json.cs
- open-a-spreadsheet-and-filter-custom-properties-by-datetime-type-then-list-them.cs
- instantiate-two-workbooks-and-copy-all-document-properties-from-source-to-destination-programmatically.cs
- load-a-template-workbook-and-clone-its-builtin-properties-into-a-newly-created-workbook.cs
- open-a-workbook-and-validate-that-documentversion-matches-a-semantic-version-pattern-before-saving.cs
- load-an-excel-file-and-verify-that-language-contains-a-valid-net-culture-code.cs
- open-a-workbook-set-scalecrop-to-true-and-ensure-the-flag-persists-when-saving-as-pdf.cs
- instantiate-a-workbook-and-deliberately-access-a-nonexistent-builtin-property-to-demonstrate-exception-handling.cs
- load-a-spreadsheet-and-use-trycatch-to-safely-read-a-custom-property-reviewer.cs
- open-a-workbook-and-confirm-that-application-metadata-fields-appear-in-saved-file-properties.cs
- load-a-workbook-modify-several-properties-and-save-the-file-to-xlsx-format-preserving-changes.cs
- open-a-workbook-update-document-properties-and-export-the-result-to-csv-format-for-downstream-processing.cs
- instantiate-a-workbook-adjust-builtin-properties-and-save-the-document-as-pdf-to-embed-metadata.cs
- batch-process-all-workbooks-in-a-folder-setting-each-files-language-property-to-en-gb.cs
- iterate-through-a-directory-of-excel-files-and-update-every-documentversion-property-to-30.cs
