using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

class ClearSpecificPivotFilter
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Region");
        worksheet.Cells["C1"].PutValue("Sales");

        string[] categories = { "Fruit", "Vegetable", "Fruit", "Vegetable", "Fruit" };
        string[] regions = { "North", "South", "East", "West", "North" };
        double[] sales = { 120, 150, 200, 130, 180 };

        for (int i = 0; i < categories.Length; i++)
        {
            worksheet.Cells[i + 1, 0].PutValue(categories[i]);
            worksheet.Cells[i + 1, 1].PutValue(regions[i]);
            worksheet.Cells[i + 1, 2].PutValue(sales[i]);
        }

        // Add a pivot table covering the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:C6", "E3", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add fields: Category as row, Region as column, Sales as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Apply a filter on the row field (Category) to show only "Fruit"
        PivotField rowField = pivotTable.RowFields[0];
        rowField.FilterByLabel(PivotFilterType.CaptionEqual, "Fruit", null);

        // Apply a filter on the column field (Region) to hide "South"
        PivotField columnField = pivotTable.ColumnFields[0];
        columnField.FilterByLabel(PivotFilterType.CaptionNotEqual, "South", null);

        // Refresh and calculate to apply the filters
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Clear filter only on the row field (field index 0) while preserving other filters
        // Using PivotFilterCollection.ClearFilter method
        pivotTable.PivotFilters.ClearFilter(0);

        // Refresh again; column filter remains active
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("ClearSpecificPivotFilter.xlsx");
    }
}