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
- verify-that-html-output-contains-only-standard-comment-tags-without-downlevel-revealed-format.cs
- test-html-files-in-legacy-browsers-to-confirm-comments-are-hidden-appropriately.cs
- document-the-configuration-required-to-turn-off-downlevel-revealed-comments.cs
- compare-page-source-before-and-after-disabling-downlevel-revealed-comments.cs
- apply-this-setting-when-exporting-workbooks-containing-extensive-cell-comments.cs
- ensure-that-comment-visibility-settings-respect-user-privacy-requirements-strictly.cs
- validate-that-disabling-downlevel-revealed-comments-does-not-affect-other-html-features.cs
- load-an-xlsx-workbook-enable-excludeunusedstyles-and-save-as-html-to-reduce-file-size.cs
- load-an-xls-workbook-set-isexportcomments-to-true-and-generate-html-that-includes-all-cell-comments.cs
- load-a-workbook-enable-exportgridlines-and-produce-html-that-displays-worksheet-gridlines.cs
- load-a-workbook-with-conditional-formatting-activate-databar-export-and-save-html-preserving-databar-visuals.cs
- combine-excludeunusedstyles-and-isexportcomments-options-then-save-workbook-as-html-for-compact-output-with-comments.cs
- process-multiple-xlsx-files-in-a-directory-applying-excludeunusedstyles-to-each-and-batch-save-reduced-size-html-files.cs
- iterate-through-all-worksheets-in-a-workbook-saving-each-as-separate-html-files-while-preserving-gridlines.cs
- validate-that-generated-html-contains-comment-tags-by-searching-for-after-enabling-isexportcomments.cs
- compare-html-file-sizes-with-and-without-excludeunusedstyles-to-quantify-reduction-achieved-by-style-omission.cs
- verify-that-unused-named-styles-are-omitted-from-the-html-output-when-excludeunusedstyles-is-enabled.cs
- load-a-workbook-with-hidden-rows-enable-exportgridlines-and-ensure-hidden-rows-do-not-affect-gridline-rendering.cs
- create-a-workbook-programmatically-add-a-comment-to-a-cell-enable-isexportcomments-and-verify-comment-appears-in-html.cs
- export-a-workbook-with-conditional-formatting-that-includes-overlapping-rules-and-ensure-both-rules-are-represented-in-html.cs
- save-a-workbook-to-html-with-exportgridlines-disabled-then-confirm-that-gridlines-are-absent-in-the-output.cs
- apply-a-custom-style-to-a-range-enable-excludeunusedstyles-and-verify-only-used-styles-appear-in-html.cs
- create-a-workbook-with-conditional-formatting-using-iconset-disable-iconset-export-and-confirm-icons-are-omitted-in-html.cs
- batch-convert-a-set-of-xls-files-to-html-applying-exportgridlines-only-to-worksheets-that-contain-tables.cs
- generate-html-from-a-workbook-while-disabling-all-conditional-formatting-export-to-produce-a-plain-visual-representation.cs
- export-a-workbook-with-conditional-formatting-that-includes-databar-and-extract-the-generated-css-classes-for-analysis.cs
- export-a-workbook-with-conditional-formatting-that-uses-threecolor-scales-and-verify-gradient-colors-appear-in-html.cs
- batch-process-workbooks-applying-both-excludeunusedstyles-and-exportgridlines-and-compare-processing-time-against-default-settings.cs
- batch-process-a-folder-of-xlsx-files-applying-excludeunusedstyles-only-to-files-larger-than-5-mb.cs
- export-a-workbook-with-conditional-formatting-that-includes-databar-and-verify-bar-lengths-match-cell-values-in-html.cs
- generate-html-from-a-workbook-while-disabling-excludeunusedstyles-to-compare-full-style-set-against-reduced-version.cs
- batch-process-a-collection-of-workbooks-applying-exportgridlines-and-isexportcomments-simultaneously-to-each-html-output.cs
- export-a-workbook-with-conditional-formatting-that-includes-colorscale-and-extract-the-css-gradient-definitions-for-reuse.cs
- export-a-workbook-with-conditional-formatting-that-includes-iconset-and-verify-icons-are-rendered-using-img-tags.cs
- batch-convert-workbooks-toggling-exportgridlines-based-on-whether-the-source-worksheet-contains-borders.cs
- export-a-workbook-with-conditional-formatting-that-includes-databar-and-generate-a-separate-css-file-containing-bar-styles.cs
- export-a-workbook-with-conditional-formatting-that-includes-colorscale-and-embed-the-scale-definitions-within-a-style-block.cs
- load-an-xlsx-workbook-and-export-to-html-using-default-settings-preserving-all-content.cs
- export-a-workbook-to-html-while-omitting-document-properties-by-setting-exportdocumentproperties-to-false.cs
- export-a-workbook-to-html-while-omitting-workbook-properties-by-disabling-exportworkbookproperties-flag.cs
- export-a-workbook-to-html-while-omitting-worksheet-properties-by-setting-exportworksheetproperties-to-false.cs
- export-a-workbook-to-html-while-omitting-document-workbook-and-worksheet-properties-simultaneously.cs
- export-a-workbook-to-html-while-omitting-workbook-and-worksheet-properties-together.cs
- export-a-workbook-to-html-while-omitting-document-and-worksheet-properties-together.cs
- export-right-to-left-aligned-text-correctly-using-default-htmlsaveoptions-during-conversion.cs
- explicitly-enable-right-to-left-text-support-by-setting-rtltextsupport-to-true-before-saving.cs
- export-only-the-defined-print-area-to-html-by-setting-exportprintareaonly-to-true.cs
- export-defined-print-area-while-omitting-document-properties-by-combining-exportprintareaonly-and-exportdocumentproperties-false.cs
- export-defined-print-area-while-omitting-worksheet-properties-by-combining-exportprintareaonly-and-exportworksheetproperties-false.cs
- export-defined-print-area-while-disabling-css-generation-by-setting-exportprintareaonly-true-and-disablecss-true.cs
- export-a-workbook-to-html-with-both-css-disabled-and-custom-properties-enabled.cs
- export-a-workbook-to-html-with-custom-properties-enabled-while-preserving-default-css-generation.cs
- disable-css-generation-during-html-export-by-setting-disablecss-to-true.cs
- enable-custom-css-properties-during-html-export-by-setting-enablecustomproperties-to-true.cs
- load-an-html-file-into-a-workbook-while-preserving-div-tag-layout-using-enabledivtaglayout.cs
- load-an-html-file-into-a-workbook-without-preserving-div-layout-by-leaving-enabledivtaglayout-false.cs
- load-html-preserving-div-layout-then-re-export-to-html-and-verify-div-structure-remains-unchanged.cs
- load-html-modify-a-cell-value-programmatically-and-export-to-html-with-default-options.cs
- load-html-modify-a-cell-and-export-to-html-while-disabling-css-generation.cs
- load-html-preserving-div-layout-modify-a-cell-and-export-to-html-while-disabling-css-generation.cs
- batch-convert-all-xlsx-files-in-a-directory-to-html-using-default-conversion-settings.cs
- batch-convert-xlsx-files-to-html-while-omitting-document-properties-for-privacy-compliance.cs
- batch-convert-xlsx-files-to-html-while-disabling-css-generation-to-reduce-file-size.cs
- export-workbook-to-html-after-defining-a-print-area-ensuring-only-that-area-appears.cs
- export-workbook-to-html-after-defining-a-print-area-and-omitting-worksheet-properties.cs
- export-workbook-to-html-after-defining-a-print-area-and-disabling-css-generation.cs
- set-htmlsaveoptionssimilarborderstyle-to-true-to-provide-fallback-borders-for-unsupported-browsers-in-html.cs
- use-htmlsaveoptionssimilarborderstyle-to-generate-fallback-border-styles-and-test-rendering-in-legacy-browsers.cs
- test-fallback-border-rendering-in-firefox-by-exporting-with-similarborderstyle-enabled-and-inspecting-css.cs
- test-fallback-border-rendering-in-safari-by-exporting-with-similarborderstyle-enabled-and-reviewing-html-output.cs
- enable-htmlsaveoptionsisiecompatible-before-saving-to-mhtml-to-allow-worksheet-tab-switching-in-internet-explorer.cs
- generate-iecompatible-mhtml-from-a-multisheet-workbook-and-verify-tab-switching-functionality-in-internet-explorer.cs
- test-mhtml-output-in-microsoft-edge-after-enabling-isiecompatible-and-assess-compatibility-with-modern-browsers.cs
- test-mhtml-output-in-google-chrome-with-isiecompatible-false-and-confirm-standard-rendering-without-iespecific-features.cs
- implement-ifilepathprovider-to-customize-html-file-paths-for-each-exported-worksheet-and-maintain-link-integrity.cs
- assign-a-custom-ifilepathprovider-instance-to-htmlsaveoptionsfilepathprovider-to-fix-broken-worksheet-hyperlinks.cs
- implement-ifilepathprovider-that-converts-worksheet-names-to-lowercase-urls-for-consistent-linking.cs
- implement-ifilepathprovider-that-generates-unique-file-names-for-worksheets-with-duplicate-titles.cs
- store-exported-html-files-in-a-temporary-folder-via-ifilepathprovider-and-clean-up-after-processing.cs
- export-workbook-to-html-with-custom-file-naming-pattern-using-ifilepathprovider-that-appends-timestamps.cs
- export-workbook-to-html-while-preserving-original-worksheet-order-by-not-altering-ifilepathprovider-mapping.cs
