// Title: Move a pivot table row item two positions forward within the same parent node using Aspose.Cells PivotItem.Move in C#
// AI Prompts: Generate C# code that uses Aspose.Cells PivotItem.Move to shift the first row field item down by two positions while keeping it under the same parent. | Show how to programmatically reorder pivot table row items by a specific offset with Aspose.Cells for .NET. | Provide an example that moves a pivot item forward by two slots in a workbook and saves the result using Aspose.Cells.
// Common Searches: aspnet move pivot row item two positions Aspose.Cells | C# Aspose.Cells PivotItem.Move example moving items within same parent | how to reorder pivot table row fields programmatically in .NET | shift pivot table item down by offset using Aspose.Cells | move first pivot item forward in Excel workbook with Aspose.Cells C#
// Tags: Aspose.Cells pivot item reorder | C# PivotItem.Move two positions | move pivot row field item within same parent | .NET pivot table item ordering | Aspose.Cells shift pivot item offset

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample creates a workbook, adds sample data, builds a pivot table with the Product field in the row area, refreshes it, then moves the first pivot row item two positions forward within the same parent node using PivotItem.Move, and finally saves the workbook as PivotItemMoveTwoPositions.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["A5"].PutValue("D");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["B4"].PutValue(300);
        sheet.Cells["B5"].PutValue(400);

        // Add a pivot table covering the data range and place it at E3
        int ptIndex = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];

        // Add the "Product" field to the row area of the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

        // Refresh and calculate the pivot table to populate items
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Get the collection of pivot items for the row field
        PivotItemCollection items = pivotTable.RowFields[0].PivotItems;

        // Move the first pivot item two positions forward within the same parent node
        // (count > 0 moves down; true indicates the same parent)
        if (items.Count > 2)
        {
            items[0].Move(2, true);
        }

        // Save the workbook with the updated pivot item order
        wb.Save("PivotItemMoveTwoPositions.xlsx");
    }
}
