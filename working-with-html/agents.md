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
