// Title: Clear a label filter on a specific pivot table row field using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that calls PivotField.ClearFilter() to remove a caption filter from a pivot table row field and then refreshes the pivot. | Provide a complete Aspose.Cells example that applies a label filter to a row field, clears it, recalculates the pivot, and saves the workbook.
// Common Searches: Aspose.Cells C# how to remove a row field label filter from a pivot table | programmatically clear a specific pivot row filter using Aspose.Cells | reset pivot table row field filter after applying caption equal filter in .NET | example of using ClearFilter on PivotField with Aspose.Cells
// Tags: Aspose.Cells pivot table row field ClearFilter | C# PivotField filter clearing | Aspose.Cells refresh pivot after filter change | Excel workbook save after pivot filter reset .NET | programmatic pivot filter management Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, builds a pivot table, applies a caption‑equal filter to the 'Category' row field, then uses PivotField.ClearFilter() to remove the filter, refreshes and recalculates the pivot, and saves the file as ClearRowFieldFilterDemo.xlsx.
class ClearPivotRowFieldFilter
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].Value = "Category";
        worksheet.Cells["B1"].Value = "Product";
        worksheet.Cells["C1"].Value = "Sales";

        worksheet.Cells["A2"].Value = "Fruit";
        worksheet.Cells["B2"].Value = "Apple";
        worksheet.Cells["C2"].Value = 120;

        worksheet.Cells["A3"].Value = "Fruit";
        worksheet.Cells["B3"].Value = "Banana";
        worksheet.Cells["C3"].Value = 80;

        worksheet.Cells["A4"].Value = "Vegetable";
        worksheet.Cells["B4"].Value = "Carrot";
        worksheet.Cells["C4"].Value = 60;

        worksheet.Cells["A5"].Value = "Vegetable";
        worksheet.Cells["B5"].Value = "Broccoli";
        worksheet.Cells["C5"].Value = 70;

        // Create a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add "Category" as a row field (field index 0)
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);

        // Add "Sales" as a data field (field index 2)
        pivotTable.AddFieldToArea(PivotFieldType.Data, 2);

        // Apply a filter on the row field to show only "Fruit" category
        PivotField rowField = pivotTable.RowFields[0];
        rowField.FilterByLabel(PivotFilterType.CaptionEqual, "Fruit", null);

        // Refresh the pivot table to apply the filter
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Clear the filter on the specific row field, restoring all items
        rowField.ClearFilter();

        // Refresh again to reflect the cleared filter
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("ClearRowFieldFilterDemo.xlsx");
    }
}
