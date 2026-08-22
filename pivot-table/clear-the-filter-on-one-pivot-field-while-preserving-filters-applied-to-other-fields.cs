// Title: How to clear a filter on a specific pivot column field while keeping other pivot filters in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that removes only the Region column filter from an Aspose.Cells pivot table without affecting the existing Category row filter. | Show how to call PivotField.ClearFilter on a single pivot field and then refresh the pivot table using Aspose.Cells. | Provide a step‑by‑step example of preserving row filters while clearing column filters in a .NET workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# clear column filter in pivot table while preserving row filter | remove specific pivot field filter Aspose.Cells .NET example | how to use PivotField.ClearFilter for a single field in Aspose.Cells | refresh pivot table after clearing a filter with Aspose.Cells C# | Aspose.Cells pivot table filter management code sample
// Tags: Aspose.Cells PivotField.ClearFilter method | clear individual pivot field filter .NET | retain row filters when clearing column filter Aspose.Cells | pivot table refresh after filter modification C# | Aspose.Cells pivot table filter handling

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

// The sample creates a workbook, adds a pivot table, applies label filters to the Category row field and Region column field, then uses PivotField.ClearFilter to remove only the column filter while preserving the row filter, refreshes the pivot table, and saves the file.
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

        worksheet.Cells["A2"].PutValue("Fruit");
        worksheet.Cells["B2"].PutValue("North");
        worksheet.Cells["C2"].PutValue(100);

        worksheet.Cells["A3"].PutValue("Fruit");
        worksheet.Cells["B3"].PutValue("South");
        worksheet.Cells["C3"].PutValue(150);

        worksheet.Cells["A4"].PutValue("Vegetable");
        worksheet.Cells["B4"].PutValue("North");
        worksheet.Cells["C4"].PutValue(200);

        worksheet.Cells["A5"].PutValue("Vegetable");
        worksheet.Cells["B5"].PutValue("South");
        worksheet.Cells["C5"].PutValue(250);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add fields: Category as row, Region as column, Sales as data
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Apply filters on both row and column fields
        // Row filter: show only "Fruit"
        pivotTable.RowFields[0].FilterByLabel(PivotFilterType.CaptionEqual, "Fruit", null);
        // Column filter: show only "North"
        pivotTable.ColumnFields[0].FilterByLabel(PivotFilterType.CaptionEqual, "North", null);

        // Refresh to apply the filters
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Clear filter only on the column field (Region) while preserving the row filter
        // Using PivotField.ClearFilter method as per the provided rule
        pivotTable.ColumnFields[0].ClearFilter();

        // Refresh again to reflect the change
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("ClearSpecificPivotFilter.xlsx");
    }
}
