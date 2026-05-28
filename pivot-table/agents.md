# Pivot Table Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Pivot Table


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Pivot Table**.

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
- create-a-new-workbook-add-source-data-and-generate-a-pivottable-on-the-first-worksheet.cs
- apply-the-compact-layout-to-an-existing-pivottable-by-invoking-the-showincompactform-method.cs
- switch-a-pivottable-to-outline-form-by-calling-the-showinoutlineform-method-on-the-table.cs
- set-a-pivottable-to-tabular-layout-using-showintabularform-and-verify-column-alignment-afterwards.cs
- assign-the-builtin-pivottablestylemedium9-autoformat-to-a-pivottable-for-quick-styling-automatically.cs
- create-a-custom-style-object-configure-its-font-color-and-apply-it-using-pivottableformat.cs
- apply-a-predefined-style-to-all-pivottable-elements-by-calling-pivottableformatall-with-the-style.cs
- clear-all-page-fields-from-a-pivottable-by-invoking-the-pivottablepagefieldsclear-method-first.cs
- remove-all-row-fields-from-a-pivottable-using-pivottablerowfieldsclear-to-reset-its-layout.cs
- set-numeric-format-000-for-a-data-field-to-display-currency-values-correctly.cs
- load-an-existing-excel-file-locate-a-pivottable-and-change-its-layout-to-tabular.cs
- add-a-calculated-field-named-profit-with-expression-revenue-cost-to-the-pivottable.cs
- hide-a-specific-pivotfield-from-the-report-area-by-setting-its-visible-property-to-false.cs
- show-a-previously-hidden-column-field-in-the-pivottable-by-toggling-its-visible-flag-back.cs
- delete-a-pivottable-from-the-worksheet-using-worksheetpivottablesremoveat-with-the-appropriate-index.cs
- iterate-through-all-pivottables-in-a-workbook-and-apply-the-outline-layout-to-each-one.cs
- batch-create-pivottables-on-multiple-worksheets-by-looping-over-data-ranges-and-calling-the-add-method.cs
- export-a-workbook-with-pivottables-to-pdf-format-while-preserving-all-pivot-formatting-details.cs
- validate-that-a-pivottable-contains-at-least-one-data-field-before-performing-any-calculations.cs
- apply-conditional-formatting-to-pivottable-values-that-exceed-a-threshold-using-the-formatcondition-feature.cs
- configure-the-pivottable-to-display-grand-totals-for-rows-only-by-setting-showrowgrandtotals-true.cs
- disable-column-grand-totals-in-a-pivottable-by-setting-showcolumngrandtotals-property-to-false.cs
- refresh-all-pivottables-in-a-workbook-after-bulk-data-changes-using-a-foreach-loop-for-each.cs
- export-a-pivottable-to-csv-by-extracting-its-data-rows-and-writing-them-with-commas.cs
- clone-an-existing-pivottable-to-a-new-worksheet-while-preserving-its-layout-and-formatting-settings.cs
- set-the-pivottables-refreshonfileopen-property-to-true-to-enable-autorefresh-when-workbook-loads.cs
- limit-the-number-of-items-displayed-in-a-pivotfield-by-adjusting-its-showitemscount-property-value.cs
- enable-drillthrough-functionality-for-a-pivottable-by-setting-the-enabledrillthrough-option-to-true.cs
- customize-the-pivottables-report-filter-caption-by-assigning-a-new-descriptive-string-to-its-caption-property.cs
- programmatically-reorder-pivotfields-within-the-row-area-to-change-the-data-hierarchy-display-order.cs
- add-a-slicer-linked-to-a-pivottable-for-interactive-filtering-using-the-slicercollection-api.cs
- remove-all-slicers-associated-with-a-pivottable-by-clearing-its-slicercollection-completely.cs
- set-the-pivottables-datacaption-property-to-a-custom-label-for-the-values-column-header.cs
- enable-repeat-item-labels-in-rows-by-setting-the-repeatitemlabels-property-to-true-on-the-pivottable.cs
- disable-automatic-grouping-of-date-fields-in-a-pivottable-by-setting-the-groupdates-property-to-false.cs
- programmatically-expand-all-collapsed-items-in-a-pivottable-using-the-expandall-method-for-full-visibility.cs
- collapse-specific-pivotitems-by-iterating-them-and-setting-each-items-isexpanded-property-to-false.cs
- set-the-pivottables-preserveformatting-flag-to-true-to-retain-applied-styles-after-each-refresh.cs
- configure-the-pivottable-to-use-the-default-data-source-connection-string-for-external-data-retrieval-operations.cs
- update-the-external-connections-command-text-for-a-pivottable-to-query-a-different-database-table.cs
- remove-all-data-fields-from-a-pivottable-to-convert-it-into-a-pure-row-and-column-summary.cs
- set-the-pivottables-displayerrorstring-property-to-a-custom-message-for-handling-calculation-errors.cs
- enable-the-showvaluesrow-property-to-display-a-separate-row-that-aggregates-values-for-better-readability.cs
- configure-the-pivottable-to-hide-empty-rows-by-setting-the-hideemptyrows-property-to-true.cs
- apply-a-theme-color-to-all-pivottable-headers-using-the-formatall-method-with-a-predefined-style.cs
- programmatically-set-the-pivottables-rowheadercaption-to-a-localized-string-for-international-user-interfaces.cs
