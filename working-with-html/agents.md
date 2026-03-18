# Working With HTML Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Working With HTML


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Working With HTML**.

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
- ensure-that-numeric-values-parsed-from-csv-files-retain-full-precision-without-c.cs
- load-a-csv-file-with-loadoptionsnumberformat-to-retain-large-numeric-values-in-f.cs
- prevent-automatic-numeric-conversion-during-csv-import-for-the-operation-by-sett.cs
- configure-loadoptionsnumberformat-to-0-to-ensure-csv-data-loads-as-plain-numeric.cs
- set-loadoptionsconvertnumericdata-to-false-when-loading-csv-files-to-disable-aut.cs
- when-loading-an-xlsx-file-configure-the-required-custom-culture-settings-to-ensu.cs
- prevent-large-numeric-values-from-being-interpreted-in-exponential-notation-duri.cs
- load-xlsx-spreadsheets-as-a-supported-input-format-for-processing-preserving-wor.cs
- load-an-xlsx-workbook-into-memory-for-programmatic-manipulation-and-data-extract.cs
- load-an-xlsx-workbook-then-export-it-to-html-preserving-worksheet-headings-in-th.cs
- load-an-xlsx-worksheet-and-save-it-as-html-enabling-htmlsaveoptionsexportheading.cs
- enable-htmlsaveoptionsexportheadings-before-converting-an-xlsx-workbook-to-html-.cs
- set-htmlsaveoptionsheadinglevel-to-define-the-resulting-heading-hierarchy-while-.cs
- load-an-xlsx-workbook-and-export-it-to-html-using-configurable-html-save-options.cs
- load-an-xlsx-workbook-and-export-its-contents-to-an-html-file-while-preserving-f.cs
- trim-unnecessary-whitespace-following-line-breaks-during-csv-import-to-ensure-cl.cs
- strip-leading-and-trailing-whitespace-from-fields-when-loading-csv-data-to-ensur.cs
- configure-load-options-to-customize-how-an-xlsx-workbook-is-imported-controlling.cs
- load-the-xlsx-workbook-and-save-it-as-html-while-disabling-downlevel-revealed-co.cs
- enable-the-htmlsaveoptionsdisabledownlevelrevealedcomments-property-when-loading.cs
- when-loading-an-xlsx-workbook-and-exporting-to-html-suppress-downlevelrevealed-c.cs
- disable-downlevel-revealed-comments-during-the-conversion-of-an-xlsx-workbook-fi.cs
- during-xlsx-to-html-conversion-omit-any-styles-not-referenced-in-the-workbook-to.cs
- convert-an-xlsx-workbook-to-html-activating-htmlsaveoptionsexcludeunusedstyles-t.cs
- enable-htmlsaveoptionsexcludeunusedstyles-during-xlsx-loading-and-html-export-to.cs
- load-an-xlsx-workbook-discard-unused-styles-and-export-the-content-to-html-forma.cs
- export-an-xlsx-worksheet-to-html-with-databar-colorscale-and-iconset-conditional.cs
- when-loading-an-xlsx-file-and-saving-as-html-set-htmlsaveoptionsexportconditiona.cs
- load-an-xlsx-workbook-and-export-it-to-html-while-retaining-all-conditional-form.cs
- when-converting-an-xlsx-workbook-to-html-include-all-cell-comments-in-the-genera.cs
- load-an-xlsx-workbook-enable-htmlsaveoptionsexportcomments-and-save-it-as-html-t.cs
- enable-comment-export-by-setting-htmlsaveoptionsexportcomments-to-true-when-load.cs
- ensure-that-cell-comments-are-retained-in-the-generated-html-when-loading-an-xls.cs
- load-an-xlsx-workbook-and-export-it-to-html-format-preserving-gridlines-in-the-r.cs
- load-an-xlsx-worksheet-and-save-it-as-html-using-htmlsaveoptionsexportgridlines-.cs
- load-an-xlsx-workbook-enable-htmlsaveoptionsexportgridlines-and-save-the-file-as.cs
- render-spreadsheet-gridlines-during-conversion-of-an-xlsx-workbook-to-html-outpu.cs
- include-document-workbook-and-worksheet-metadata-during-xlsx-to-html-conversion-.cs
- convert-an-xlsx-workbook-to-html-enabling-htmlsaveoptionsexportdocumentpropertie.cs
- enable-htmlsaveoptionsexportdocumentproperties-when-converting-an-xlsx-workbook-.cs
- load-an-xlsx-workbook-configure-html-export-options-to-embed-metadata-and-save-t.cs
- load-an-html-file-into-a-workbook-enable-htmlloadoptionspreservedivlayout-and-sa.cs
- enable-preservedivlayout-on-htmlloadoptions-when-importing-html-then-export-the-.cs
- preserve-div-tag-layout-when-importing-html-into-an-excel-workbook-and-exporting.cs
- ensure-div-layout-is-retained-when-importing-html-content-and-exporting-to-xlsx-.cs
- configure-html-load-options-to-import-html-content-and-export-it-as-an-xlsx-work.cs
- load-an-html-document-and-convert-it-to-an-xlsx-workbook-while-preserving-cell-d.cs
- export-an-xlsx-workbook-to-html-preserving-righttoleft-text-expansion-during-the.cs
- export-the-xlsx-worksheet-with-righttoleft-text-to-html-enabling-the-expandrtlte.cs
- load-an-xlsx-workbook-enable-htmlsaveoptionsexpandrtltext-and-export-the-file-co.cs
- process-an-xlsx-workbook-containing-rtl-text-and-export-it-to-html-ensuring-prop.cs
- load-an-xlsx-workbook-extract-its-defined-print-area-and-generate-an-html-repres.cs
- export-the-worksheets-defined-print-area-to-html-by-setting-htmlsaveoptionsexpor.cs
- configure-htmlsaveoptionsexportprintareaonly-to-true-when-loading-an-xlsx-workbo.cs
- load-an-xlsx-workbook-specify-a-print-area-and-export-that-region-to-html-format.cs
- when-converting-an-xlsx-to-html-substitute-unsupported-cell-border-styles-with-v.cs
- export-a-loaded-xlsx-worksheet-to-html-using-htmlsaveoptionssimulateunsupportedb.cs
- enable-htmlsaveoptionssimulateunsupportedborders-when-converting-an-xlsx-workboo.cs
- approximate-unsupported-cell-borders-when-converting-an-xlsx-workbook-to-html-ou.cs
- load-an-xlsx-workbook-and-export-it-to-an-iecompatible-mhtml-file-preserving-for.cs
- load-an-xlsx-workbook-and-save-it-as-mhtml-enabling-internet-explorer-compatibil.cs
- load-an-xlsx-workbook-and-export-it-as-an-iecompatible-mhtml-file-while-preservi.cs
- load-an-xlsx-file-and-configure-save-options-to-export-the-workbook-as-an-mhtml-.cs
- programmatically-load-an-xlsx-workbook-then-convert-and-save-it-as-an-mhtml-docu.cs
- implement-ifilepathprovider-to-supply-the-html-output-path-when-loading-an-xlsx-.cs
- implement-ifilepathprovider-load-an-xlsx-workbook-and-use-htmlsaveoptions-to-exp.cs
- develop-a-class-implementing-ifilepathprovider-overriding-getfilepath-to-load-xl.cs
- specify-custom-output-locations-when-loading-an-xlsx-workbook-and-exporting-it-a.cs
