# Working With Tables Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Working With Tables


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Working With Tables**.

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
- apply-a-predefined-table-style-to-the-created-table-and-preserve-the-original-formatting.cs
- add-a-totals-row-to-the-table-and-configure-sum-formulas-for-numeric-columns.cs
- set-a-custom-formula-in-the-totals-row-to-calculate-average-of-a-specific-column.cs
- hide-the-table-header-row-while-keeping-the-data-rows-visible-for-reporting-purposes.cs
- enable-autofilter-on-the-table-and-define-a-filter-to-show-only-rows-with-values-above-threshold.cs
- group-rows-within-the-table-based-on-category-column-and-collapse-the-groups-for-compact-view.cs
- protect-the-entire-table-with-a-password-allowing-only-readonly-access-for-external-users.cs
- lock-specific-columns-in-the-table-to-prevent-accidental-modification-while-allowing-other-columns-to-edit.cs
- add-a-comment-to-the-table-object-describing-its-purpose-and-retrieve-the-comment-text-programmatically.cs
- convert-the-existing-list-object-into-a-structured-table-to-leverage-advanced-table-features.cs
- disable-the-ability-for-users-to-add-new-rows-to-the-list-object-to-enforce-fixed-dataset-size.cs
- set-the-list-object-to-display-a-totals-row-and-configure-count-aggregation-for-a-text-column.cs
- load-an-existing-excel-workbook-containing-query-tables-and-enumerate-all-tables-linked-to-external-data-sources.cs
- read-the-metadata-of-a-query-table-including-connection-string-command-type-and-refresh-interval.cs
- write-a-datatable-object-into-a-new-worksheet-table-mapping-column-names-to-table-headers-automatically.cs
- apply-conditional-formatting-to-a-table-column-that-highlights-cells-exceeding-a-defined-numeric-threshold.cs
- autofit-all-columns-of-a-table-to-match-the-longest-cell-content-for-optimal-display.cs
- set-a-custom-column-width-for-a-specific-table-column-to-accommodate-long-text-strings.cs
- validate-that-a-table-contains-no-duplicate-rows-based-on-a-combination-of-key-columns.cs
- insert-a-new-row-into-a-table-and-automatically-copy-the-formatting-from-the-previous-row.cs
- apply-a-unique-index-to-a-table-column-to-enforce-data-uniqueness-during-data-entry.cs
- enable-the-tables-autoexpand-feature-so-that-adding-data-below-expands-the-table-range-automatically.cs
- apply-a-builtin-table-style-that-matches-the-workbooks-theme-for-consistent-visual-appearance.cs
- add-a-slicer-linked-to-a-table-column-to-provide-interactive-filtering-in-the-worksheet.cs
- remove-an-existing-slicer-from-a-table-and-clean-up-associated-connections.cs
- create-a-data-validation-list-that-pulls-its-items-directly-from-a-column-in-a-worksheet-table.cs
- set-the-tables-show-header-row-option-to-false-for-a-compact-layout-in-a-dashboard-view.cs
- load-a-workbook-locate-a-table-by-name-and-export-its-contents-to-an-html-fragment.cs
- generate-a-pivot-table-based-on-an-existing-worksheet-table-and-place-it-on-a-new-worksheet.cs
- refresh-all-pivot-tables-that-reference-a-specific-worksheet-table-after-updating-its-underlying-data.cs
- use-structured-references-in-formulas-that-refer-to-table-columns-ensuring-automatic-range-adjustments.cs
- hide-a-table-column-programmatically-and-later-unhide-it-based-on-user-interaction-criteria.cs
- create-a-duplicate-of-an-existing-table-on-another-worksheet-while-preserving-its-style-and-formulas.cs
- move-a-table-to-a-different-position-within-the-same-worksheet-updating-all-structured-references-automatically.cs
- delete-a-specific-row-from-a-table-using-its-primary-key-value-to-locate-the-target.cs
- export-a-worksheet-table-to-a-json-string-preserving-column-names-as-json-object-keys.cs
- calculate-a-running-total-column-within-a-table-using-a-formula-that-references-previous-rows.cs
- set-the-tables-show-totals-row-option-to-false-removing-the-totals-row-from-the-display.cs
- disable-autoexpand-for-a-table-to-keep-its-range-fixed-despite-additional-rows-being-entered.cs
- change-the-table-style-to-a-custom-xmldefined-style-to-meet-corporate-branding-guidelines.cs
- programmatically-retrieve-the-address-range-of-a-table-and-use-it-as-a-named-range-for-formulas.cs
- update-the-named-range-that-references-a-table-after-expanding-the-table-to-include-new-rows.cs
- enable-the-tables-show-header-row-option-and-customize-the-header-font-color-for-emphasis.cs
- add-a-new-column-to-an-existing-table-and-set-its-default-value-using-a-constant-expression.cs
- remove-an-unwanted-column-from-a-table-while-preserving-the-data-in-other-columns.cs
- reorder-columns-in-a-table-to-match-a-predefined-layout-required-by-downstream-processing-scripts.cs
- apply-a-filter-that-selects-rows-where-the-status-column-equals-completed-and-hide-the-rest.cs
- clear-all-filters-applied-to-a-table-restoring-the-full-dataset-visibility-for-analysis.cs
