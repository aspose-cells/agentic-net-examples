# Cells Data Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Cells Data


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Cells Data**.

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
- load-an-excel-workbook-from-a-file-path-and-verify-it-opens-successfully.cs
- select-a-worksheet-by-name-and-obtain-its-cells-collection-for-further-operations.cs
- access-cell-b2-using-its-a1-style-name-set-a-numeric-value-and-save-the-workbook.cs
- access-a-cell-using-zerobased-row-and-column-indices-read-its-value-and-log-it.cs
- access-a-cell-by-its-numeric-index-within-the-cells-collection-modify-its-background-color-and-save.cs
- retrieve-the-worksheets-maximum-display-range-and-use-it-to-define-a-print-area.cs
- iterate-through-all-cells-in-the-maximum-display-range-to-count-nonempty-cells.cs
- convert-all-stringbased-numeric-values-in-the-entire-workbook-to-true-numbers-using-convertstringtonumericvalue.cs
- convert-string-numeric-values-only-in-the-first-worksheet-while-leaving-other-sheets-unchanged.cs
- convert-numeric-strings-within-range-a1c10-on-a-worksheet-and-verify-conversion.cs
- convert-numeric-strings-in-a-range-using-culturespecific-decimal-separators-and-ensure-correct-parsing.cs
- add-subtotals-summing-column-d-for-rows-2100-placing-results-at-each-groups-bottom.cs
- add-subtotals-that-count-entries-in-column-b-for-rows-5200-using-outline-grouping.cs
- add-subtotals-that-calculate-the-average-of-column-f-positioning-the-summary-rows-at-the-top.cs
- add-subtotals-using-the-max-function-on-column-c-with-summary-rows-placed-after-each-group.cs
- add-subtotals-using-the-stddev-function-on-column-h-and-set-summary-position-to-bottom.cs
- add-subtotals-using-the-var-function-on-column-i-with-outline-enabled-for-hierarchical-display.cs
- add-subtotals-using-the-product-function-on-column-j-placing-summary-rows-at-the-top-of-groups.cs
- add-subtotals-using-the-countnumbers-function-on-column-k-without-creating-outline-levels.cs
- add-subtotals-for-multiple-columns-simultaneously-summing-columns-m-and-n-together.cs
- apply-subtotals-after-sorting-the-worksheet-by-column-a-to-ensure-grouped-data-is-correct.cs
- apply-subtotals-on-filtered-data-confirming-that-only-visible-rows-are-included-in-calculations.cs
- apply-subtotals-with-isoutline-set-to-false-producing-a-flat-list-of-summary-rows.cs
- apply-subtotals-with-summaryposition-set-to-top-inserting-summary-rows-before-each-group.cs
- apply-subtotals-with-summaryposition-set-to-bottom-inserting-summary-rows-after-each-group.cs
- use-cellssubtotal-with-column-index-derived-from-column-name-sales-to-sum-sales-figures.cs
- use-cellssubtotal-with-startrow-set-to-zero-handling-zerobased-indexing-correctly.cs
- use-cellssubtotal-with-endrow-set-to-the-last-used-row-ensuring-full-range-coverage.cs
- handle-an-outofrange-column-index-in-cellssubtotal-gracefully-by-catching-the-exception.cs
- validate-that-cellssubtotal-throws-an-informative-error-when-startrow-exceeds-endrow.cs
- clear-the-autofilter-criteria-on-column-aj-and-display-all-rows-again.cs
- save-the-modified-workbook-to-an-xlsx-file-while-preserving-original-formatting.cs
- save-the-workbook-to-a-csv-file-ensuring-numeric-values-are-not-quoted-unnecessarily.cs
- search-for-cells-containing-the-text-pending-and-highlight-them-with-yellow-fill.cs
- find-all-numeric-values-greater-than-1000-across-the-worksheet-and-collect-their-addresses.cs
- perform-caseinsensitive-search-for-total-revenue-and-replace-it-with-revenue-total.cs
- locate-cells-with-formulas-that-return-div0-errors-and-replace-them-with-zero.cs
- identify-duplicate-entries-in-column-q-and-mark-them-with-red-background.cs
- search-for-cells-containing-the-word-error-ignoring-case-and-highlight-them-orange.cs
- find-cells-with-text-longer-than-50-characters-and-truncate-them-to-50-characters.cs
- locate-cells-with-leading-or-trailing-spaces-and-trim-the-whitespace-programmatically.cs
- find-cells-with-formulas-using-volatile-functions-and-list-their-addresses-for-review.cs
- sort-worksheet-rows-by-column-a-in-ascending-order-using-a-datasorter-instance.cs
- apply-descending-sort-on-column-d-while-preserving-original-row-grouping-via-addkey-method.cs
- sort-data-based-on-cell-background-colors-in-column-b-using-sortkey-cellcolor-property.cs
- create-a-custom-sort-list-for-months-and-sort-column-c-according-to-that-list.cs
- perform-multilevel-sorting-first-column-e-ascending-then-column-f-descending.cs
- configure-datasorter-to-ignore-hidden-rows-while-sorting-column-l-in-descending-order.cs
- set-custom-sort-order-for-priority-levels-high-medium-low-and-sort-column-m.cs
- apply-backgroundcolor-sorting-on-column-u-treating-empty-cells-as-lowest-priority.cs
- sort-rows-based-on-a-computed-helper-column-that-concatenates-first-and-last-names.cs
- perform-caseinsensitive-sort-on-column-af-while-treating-numeric-strings-as-numbers.cs
- create-a-datasorter-instance-with-stable-sorting-enabled-to-maintain-relative-order-of-equal-keys.cs
- sort-data-while-preserving-merged-cells-by-disabling-merge-handling-in-datasorter-options.cs
- define-a-cellarea-covering-rows-2100-and-apply-wholenumber-validation-to-that-range.cs
- add-validation-restricting-column-g-values-to-integers-between-10-and-500.cs
- enable-incell-dropdown-for-column-h-by-setting-validationincelldropdown-to-true.cs
- retrieve-validation-details-of-cell-j5-and-log-its-type-and-formula-values.cs
- check-whether-cell-k10-uses-an-incell-dropdown-and-output-the-result-to-console.cs
- add-validation-to-a-dynamic-range-that-expands-as-new-rows-are-inserted-using-cellarea.cs
- create-a-validation-allowing-dates-between-01012020-and-12312025-in-column-n.cs
- configure-validation-to-show-input-message-prompting-users-to-enter-a-valid-email-address.cs
- add-validation-that-disallows-blank-entries-in-column-v-and-displays-an-error-alert.cs
- validate-that-column-ak-contains-unique-email-addresses-and-display-an-error-for-duplicates.cs
- create-a-validation-that-only-permits-time-values-between-0900-and-1700-in-column-t.cs
- add-validation-restricting-column-s-values-to-a-predefined-array-of-strings.cs
- export-worksheet-rows-150-and-columns-ad-to-a-csv-string.cs
- export-worksheet-data-to-json-format-using-custom-serialization-of-cell-values-and-styles.cs
- export-worksheet-data-to-xml-file-preserving-cell-data-types-and-formatting-attributes.cs
- export-worksheet-data-to-pdf-while-preserving-cell-background-colors-and-borders.cs
- export-worksheet-data-to-an-html-file-with-embedded-css-to-retain-cell-styling.cs
- export-worksheet-data-to-a-markdown-table-preserving-header-formatting-and-alignment.cs
- export-only-validated-cells-from-a-worksheet-to-a-json-array-for-downstream-processing.cs
- export-worksheet-rows-that-fail-validation-to-a-separate-sheet-for-error-analysis.cs
- export-worksheet-data-to-a-tabdelimited-text-file-preserving-numeric-formatting.cs
- export-worksheet-data-to-a-fixedwidth-text-file-using-custom-column-width-definitions.cs
- retrieve-a-cells-formatted-string-value-using-getstringvalue-with-the-withformatting-strategy.cs
- obtain-a-cells-raw-numeric-string-by-calling-getstringvalue-with-the-withoutformatting-option.cs
- assign-simple-html-markup-to-a-cells-htmlstring-property-to-display-bold-and-italic-text.cs
- import-a-twodimensional-double-array-into-a-worksheet-starting-at-row-five-and-column-two.cs
- load-a-onedimensional-string-array-into-cells-beginning-at-the-first-row-and-first-column.cs
- use-importarraylist-to-add-values-from-an-arraylist-into-a-worksheet-beginning-at-row-three.cs
- import-a-collection-of-custom-objects-mapping-properties-to-columns-starting-at-row-two-column-one.cs
- enable-preservehtml-option-in-importtableoptions-to-keep-html-formatting-when-importing-rich-text-data.cs
- set-checkmergedcells-to-true-before-importing-custom-objects-to-correctly-populate-merged-cell-ranges.cs
- adjust-the-firstrow-parameter-to-shift-existing-rows-down-before-inserting-a-new-data-table.cs
- validate-cell-content-by-retrieving-raw-string-values-and-comparing-them-against-expected-numeric-strings.cs
- combine-importarray-and-importcustomobjects-calls-to-populate-a-worksheet-from-heterogeneous-data-sources-in-one-workflow.cs
- insert-a-hyperlink-into-a-cell-using-the-hyperlinkcollectionadd-method-with-display-text-and-url.cs
- create-a-new-worksheet-import-an-array-of-dates-and-format-cells-to-display-short-date-pattern.cs
- read-a-cells-formatted-string-replace-placeholder-tokens-and-write-the-updated-string-back-to-the-cell.cs
- import-data-from-a-csv-file-into-a-worksheet-using-importdata-with-a-custom-icellsdatatable-implementation.cs
- set-the-firstrow-offset-to-zero-to-overwrite-existing-data-when-importing-a-new-dataset.cs
- retrieve-a-cells-formatted-string-parse-currency-symbols-and-store-numeric-value-in-adjacent-cell-for-calculations.cs
- import-data-from-an-arraylist-containing-mixed-types-specifying-column-data-types-to-ensure-correct-cell-formatting.cs
- create-a-workbook-add-a-worksheet-and-import-custom-objects-while-preserving-html-formatting-in-description-fields.cs
- after-importing-data-iterate-through-cells-to-count-how-many-contain-nonempty-html-content.cs
- use-getstringvalue-with-withformatting-to-generate-a-userfriendly-report-line-that-includes-currency-symbols.cs
- import-a-twodimensional-array-of-booleans-and-set-cell-style-to-display-checkmarks-for-true-values.cs
- create-a-hyperlink-that-references-a-cell-in-another-worksheet-using-the-internal-excel-address-format.cs
- retrieve-raw-strings-from-a-cell-range-concatenate-them-and-write-the-result-to-a-summary-cell.cs
- import-an-array-of-timestamps-then-format-cells-to-display-time-in-hhmmss-format-for-readability.cs
- use-importcustomobjects-with-a-mapping-dictionary-to-rename-columns-during-data-import-from-objects.cs
- set-importtableoptionscheckmergedcells-to-false-to-intentionally-skip-writing-into-merged-cell-ranges.cs
