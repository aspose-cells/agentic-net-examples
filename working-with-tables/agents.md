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
- create-a-new-worksheet-table-from-a-range-of-cells-and-assign-a-custom-name.cs
- apply-a-predefined-table-style-to-the-created-table-and-preserve-the-original-formatting.cs
- add-a-totals-row-to-the-table-and-configure-sum-formulas-for-numeric-columns.cs
- set-a-custom-formula-in-the-totals-row-to-calculate-average-of-a-specific-column.cs
- hide-the-table-header-row-while-keeping-the-data-rows-visible-for-reporting-purposes.cs
- enable-autofilter-on-the-table-and-define-a-filter-to-show-only-rows-with-values-above-threshold.cs
- sort-the-table-by-two-columns-first-ascending-by-date-then-descending-by-amount.cs
- group-rows-within-the-table-based-on-category-column-and-collapse-the-groups-for-compact-view.cs
- protect-the-entire-table-with-a-password-allowing-only-readonly-access-for-external-users.cs
- unprotect-the-previously-secured-table-using-the-correct-password-to-enable-editing-operations.cs
- lock-specific-columns-in-the-table-to-prevent-accidental-modification-while-allowing-other-columns-to-edit.cs
- add-a-comment-to-the-table-object-describing-its-purpose-and-retrieve-the-comment-text-programmatically.cs
- update-the-existing-table-comment-to-include-version-information-and-author-initials-for-documentation-tracking.cs
- delete-the-comment-attached-to-the-table-to-clean-up-metadata-after-final-review.cs
- convert-the-existing-list-object-into-a-structured-table-to-leverage-advanced-table-features.cs
- create-a-list-object-from-a-dynamic-range-and-enable-automatic-expansion-when-new-rows-are-added.cs
- disable-the-ability-for-users-to-add-new-rows-to-the-list-object-to-enforce-fixed-dataset-size.cs
- enable-the-list-objects-header-row-and-customize-its-background-color-using-a-predefined-style.cs
- set-the-list-object-to-display-a-totals-row-and-configure-count-aggregation-for-a-text-column.cs
- load-an-existing-excel-workbook-containing-query-tables-and-enumerate-all-tables-linked-to-external-data-sources.cs
- set-the-background-refresh-property-of-a-query-table-to-false-ensuring-synchronous-data-retrieval.cs
- export-the-data-from-a-query-table-to-a-csv-file-while-preserving-column-headers-and-data-types.cs
- read-the-metadata-of-a-query-table-including-connection-string-command-type-and-refresh-interval.cs
- write-a-datatable-object-into-a-new-worksheet-table-mapping-column-names-to-table-headers-automatically.cs
- load-a-workbook-locate-a-table-by-name-and-export-its-contents-to-an-html-fragment.cs
- create-a-chart-that-uses-a-worksheet-table-as-its-data-source-and-apply-a-predefined-chart-style.cs
- refresh-all-pivot-tables-that-reference-a-specific-worksheet-table-after-updating-its-underlying-data.cs
