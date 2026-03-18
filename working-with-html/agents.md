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
- load-an-xlsx-workbook-into-memory-and-export-it-as-an-html-document-while-preser.cs
- when-converting-an-xlsx-worksheet-to-html-generate-the-stylesheet-as-a-distinct-.cs
- export-a-worksheet-from-an-xlsx-file-to-html-storing-css-in-separate-files-by-en.cs
- configure-htmlsaveoptions-to-export-css-separately-load-an-xlsx-workbook-and-sav.cs
- load-an-xlsx-file-and-export-it-to-html-extracting-styles-into-an-external-css-s.cs
- load-an-html-document-and-generate-an-equivalent-pdf-file-while-preserving-its-l.cs
- load-an-html-file-using-htmlloadoptions-then-export-the-workbook-and-generate-a-.cs
- load-an-html-document-using-htmlloadoptions-and-convert-it-directly-to-a-pdf-fil.cs
- instantiate-a-workbook-load-html-content-then-export-it-as-a-pdf-using-the-appro.cs
- load-an-html-document-into-memory-and-export-it-directly-as-a-pdf-file-while-pre.cs
- use-a-streamprovider-to-load-html-content-and-generate-an-xlsx-workbook-then-sav.cs
- implement-a-custom-istreamprovider-to-load-an-html-document-and-export-it-direct.cs
- create-an-istreamprovider-implementation-that-handles-html-input-loading-and-xls.cs
- load-an-html-document-into-the-library-then-export-the-content-as-an-xlsx-workbo.cs
- load-an-xlsx-workbook-and-export-it-to-html-by-implementing-a-custom-istreamprov.cs
- create-an-istreamprovider-implementation-to-handle-output-streams-for-loading-xl.cs
- apply-gradient-fill-rendering-to-wordart-objects-when-converting-xlsx-spreadshee.cs
- ensure-wordart-shapes-retain-their-gradient-fill-when-converting-an-xlsx-workboo.cs
- export-a-spreadsheet-worksheet-to-html-with-overlaid-objects-hidden-by-activatin.cs
- enable-the-crosshideright-option-when-converting-an-xlsx-workbook-to-html-by-set.cs
- load-an-xlsx-workbook-apply-crosshideright-to-conceal-overlaid-cells-and-export-.cs
- generate-html-output-from-an-xlsx-workbook-while-suppressing-any-overlaid-elemen.cs
- load-an-xlsx-worksheet-and-export-it-to-html-configuring-htmlsaveoptionstablecss.cs
- configure-htmlsaveoptions-to-assign-the-tablecssid-value-myprefix_-when-converti.cs
- when-converting-an-xlsx-workbook-to-html-apply-the-htmlsaveoptionstablecssid-pre.cs
- load-an-xlsx-workbook-and-prepend-a-specified-prefix-to-generated-html-table-css.cs
- load-an-html-document-into-a-workbook-enable-column-and-row-autofit-via-htmlload.cs
- configure-htmlloadoptionsautofitcolumns-to-true-when-loading-html-then-export-as.cs
- enable-automatic-row-fitting-when-loading-html-by-setting-htmlloadoptionsautofit.cs
- load-an-html-file-into-a-workbook-autofit-all-columns-and-rows-and-export-the-re.cs
- apply-autofit-to-columns-when-importing-html-content-and-exporting-the-workbook-.cs
- load-an-html-file-into-a-workbook-enable-htmlloadoptionsrecognizeselfclosingtags.cs
- enable-selfclosing-tag-recognition-in-htmlloadoptions-then-load-html-content-and.cs
- identify-and-correctly-handle-selfclosing-html-tags-during-html-import-and-subse.cs
- parse-html-containing-selfclosing-tags-then-export-the-content-to-an-xlsx-workbo.cs
- export-an-xlsx-worksheet-to-html-configuring-column-widths-in-em-or-percent-via-.cs
- configure-htmlsaveoptionscolumnwidthunit-to-em-or-percent-when-loading-an-xlsx-w.cs
- configure-column-widths-in-loaded-xlsx-files-to-use-scalable-units-such-as-em-or.cs
- load-an-xlsx-file-apply-column-widths-using-scalable-units-and-export-the-workbo.cs
- convert-an-xlsx-worksheet-to-html-while-rendering-strikethrough-text-by-appropri.cs
- configure-htmlsaveoptionscrossstringtype-to-htmlcrosstypestrikethrough-when-load.cs
- configure-htmlcrosstype-to-apply-string-crossing-when-converting-a-loaded-xlsx-w.cs
- specify-cross-string-handling-behavior-when-loading-an-xlsx-workbook-and-saving-.cs
- load-an-xlsx-workbook-and-export-it-as-html-preserving-cell-tooltips-in-the-resu.cs
- load-the-xlsx-workbook-and-save-it-as-html-with-htmlsaveoptionsexportcommentsast.cs
- when-converting-an-xlsx-workbook-to-html-enable-htmlsaveoptionsexportcommentsast.cs
- include-cell-tooltip-information-in-the-html-generated-when-converting-an-xlsx-w.cs
- when-converting-an-xlsx-workbook-to-html-ensure-hidden-worksheet-elements-are-ex.cs
- load-an-xlsx-workbook-and-export-it-to-html-with-htmlsaveoptionsexporthiddenelem.cs
- load-an-xlsx-workbook-configure-export-settings-to-omit-hidden-rows-columns-and-.cs
- transform-an-xlsx-workbook-into-html-using-htmlsaveoptions-with-presentationpref.cs
- load-an-xlsx-workbook-configure-htmlsaveoptionspresentationpreference-to-bestfit.cs
- load-an-xlsx-file-and-export-it-to-html-applying-the-presentationpreference-sett.cs
- apply-presentationpreference-settings-when-loading-an-xlsx-workbook-and-exportin.cs
- export-an-xlsx-worksheet-to-html-configuring-htmlsaveoptionslinktarget_blank-so-.cs
- load-an-xlsx-workbook-configure-htmlsaveoptions-with-linktarget-set-to-_blank-an.cs
- alter-the-html-hyperlink-target-type-when-converting-an-xlsx-workbook-to-html-ou.cs
- specify-the-hyperlink-target-when-converting-a-loaded-xlsx-workbook-to-an-html-f.cs
- load-an-xlsx-workbook-and-export-it-to-html-with-all-css-disabled-via-htmlsaveop.cs
- configure-htmlsaveoptionsdisablecss-to-true-when-converting-an-xlsx-workbook-to-.cs
- load-an-xlsx-workbook-and-save-it-as-html-without-generating-any-css-styles.cs
- load-an-xlsx-workbook-and-save-it-as-html-without-explicitly-including-css-style.cs
- load-an-xlsx-workbook-and-save-it-as-html-with-css-custom-properties-enabled-via.cs
- when-converting-an-xlsx-workbook-to-html-configure-htmlsaveoptions-to-enable-css.cs
- enable-css-custom-properties-during-conversion-of-an-xlsx-workbook-to-html-prese.cs
