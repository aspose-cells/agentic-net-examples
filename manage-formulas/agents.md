# Manage formulas Examples

This folder contains **Aspose.Cells for .NET** code examples related to:

Manage formulas


## Purpose

These examples demonstrate common **Aspose.Cells APIs** used when working with:

- Workbooks
- Worksheets
- Cells
- Formulas
- Charts
- Data operations


## Example Files

Each `.cs` file demonstrates a specific task related to **Manage formulas**.

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
- change-source-data-for-a-filter-dynamic-array-formula-then-recalculate-workbook-to-update-results.cs
- retrieve-the-spilled-range-address-of-a-dynamic-array-formula-located-in-cell-c3-programmatically.cs
- programmatically-clear-the-spilled-range-of-a-dynamic-array-formula-without-deleting-the-original-formula-cell.cs
- create-a-dynamic-array-formula-that-spills-into-empty-rows-then-insert-data-to-shift-the-spill-range.cs
- create-a-dynamic-array-formula-that-references-a-table-column-then-delete-the-table-and-observe-formula-error.cs
- create-a-dynamic-array-formula-that-references-a-spill-then-use-it-in-a-sum-formula-on-another-sheet.cs
- create-a-listobject-named-salestable-add-a-column-with-a-sum-formula-and-test-propagation.cs
- insert-a-new-row-into-salestable-and-confirm-the-column-formula-automatically-calculates-for-the-new-entry.cs
- update-the-formula-of-a-table-column-to-include-if-logic-then-add-rows-to-verify-new-behavior.cs
- remove-a-column-from-a-listobject-and-ensure-its-associated-formula-no-longer-appears-in-subsequent-rows.cs
- convert-a-listobject-back-to-a-regular-range-preserving-existing-formulas-within-the-cells.cs
- add-a-calculated-column-to-a-listobject-using-the-xlookup-function-and-verify-automatic-propagation.cs
- define-a-named-range-called-dataset-covering-a1c10-then-rename-it-to-reportdata-using-nametext.cs
- replace-an-existing-named-range-with-a-larger-area-using-the-namerefersto-property-and-recalculate-formulas.cs
- create-a-named-range-that-references-a-dynamic-array-spill-and-use-it-in-subsequent-formulas.cs
- create-a-composite-named-range-by-unioning-three-separate-ranges-and-assign-a-custom-style-to-the-result.cs
- create-two-separate-range-objects-perform-union-operation-and-iterate-through-the-resulting-collection.cs
- identify-overlapping-cells-between-range-a5b15-and-range-b10c20-using-intersect-method.cs
- detect-intersecting-area-between-two-named-ranges-then-highlight-the-intersected-cells-with-yellow-fill.cs
- create-a-style-object-set-solid-fill-to-light-blue-bold-font-and-apply-to-dataset-range.cs
- apply-a-background-color-to-the-intersected-area-of-two-named-ranges-and-save-the-workbook-as-xlsx.cs
- clear-contents-of-the-named-range-reportdata-without-deleting-the-range-definition-itself.cs
- remove-the-named-range-summarydata-from-the-workbook-and-verify-it-no-longer-appears-in-the-collection.cs
