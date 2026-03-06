using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

class ClearPivotFieldFilters
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].Value = "Category";
        worksheet.Cells["B1"].Value = "Sales";
        worksheet.Cells["A2"].Value = "Fruit";
        worksheet.Cells["B2"].Value = 100;
        worksheet.Cells["A3"].Value = "Vegetable";
        worksheet.Cells["B3"].Value = 150;
        worksheet.Cells["A4"].Value = "Fruit";
        worksheet.Cells["B4"].Value = 200;
        worksheet.Cells["A5"].Value = "Vegetable";
        worksheet.Cells["B5"].Value = 250;

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add row and data fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the row field (Category) to work with its filters
        PivotField categoryField = pivotTable.RowFields[0];

        // Apply a filter to demonstrate that it can be cleared later
        categoryField.FilterByLabel(PivotFilterType.CaptionEqual, "Fruit", null);

        // Clear all filters from the pivot field using ClearFilter()
        categoryField.ClearFilter();

        // Refresh the pivot table to reflect the removal of filters
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("ClearPivotFieldFilters.xlsx");
    }
}