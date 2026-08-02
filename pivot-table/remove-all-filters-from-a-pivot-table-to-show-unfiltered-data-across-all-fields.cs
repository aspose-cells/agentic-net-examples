using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

class RemoveAllPivotFilters
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("Fruit");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Vegetable");
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["A4"].PutValue("Fruit");
        worksheet.Cells["B4"].PutValue(150);
        worksheet.Cells["A5"].PutValue("Vegetable");
        worksheet.Cells["B5"].PutValue(300);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Apply a filter to demonstrate clearing it later
        PivotField rowField = pivotTable.RowFields[0];
        rowField.FilterByLabel(PivotFilterType.CaptionEqual, "Fruit", null);

        // Refresh to apply the filter
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // ---------- Remove all filters from the pivot table ----------
        // Clear filters from row fields
        foreach (PivotField pf in pivotTable.RowFields)
        {
            pf.ClearFilter();
        }

        // Clear filters from column fields (if any)
        foreach (PivotField pf in pivotTable.ColumnFields)
        {
            pf.ClearFilter();
        }

        // Clear filters from page fields (if any)
        foreach (PivotField pf in pivotTable.PageFields)
        {
            pf.ClearFilter();
        }

        // Clear filters from data fields (generally not filtered, but included for completeness)
        foreach (PivotField pf in pivotTable.DataFields)
        {
            pf.ClearFilter();
        }

        // Refresh the pivot table to reflect the removal of all filters
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the unfiltered pivot table
        workbook.Save("PivotTable_NoFilters.xlsx");
    }
}