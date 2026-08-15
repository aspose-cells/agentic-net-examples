// Title: Move a Pivot Table Row Item Two Positions Forward with Aspose.Cells PivotItem.Move (C#)
// Description: This example creates a workbook, adds sample data, builds a pivot table, places the "Product" field in the row area, refreshes and calculates the pivot, then moves the first row item two positions down within the same parent node using PivotItem.Move(2, true). The pivot is recalculated and saved as an XLSX file.
// Keywords: Aspose.Cells PivotItem.Move | C# pivot table reorder | move pivot item two positions | isSameParent flag | programmatic pivot item sorting | Aspose.Cells pivot table example
// Common Searches: Aspose.Cells move pivot item down two positions | PivotItem.Move same parent C# | reorder row items in Aspose.Cells pivot table | how to shift pivot items programmatically | Aspose.Cells pivot table item ordering
// Developer Intent: Programmatically shift a pivot table row item two positions forward while keeping it under the same parent node.
// Use Cases: Adjust the display order of row items after data changes without manual sorting. | Implement custom sorting rules by moving specific items to a desired position. | Prepare a pivot table with a predefined item sequence before exporting or sharing the workbook.
// AI Prompts: Show how to move a pivot item up one position using Aspose.Cells PivotItem.Move. | Explain the purpose of the isSameParent parameter when reordering pivot items. | Provide code to move multiple pivot items based on a condition in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example creates a workbook, adds sample data, builds a pivot table, places the "Product" field in the row area, refreshes and calculates the pivot, then moves the first row item two positions down within the same parent node using PivotItem.Move(2, true). The pivot is recalculated and saved as an XLSX file.
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
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["B4"].PutValue(300);

        // Add a pivot table covering the data range and place it at E3
        int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];

        // Add the "Product" field to the row area of the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");

        // Refresh and calculate to populate the pivot items
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Get the collection of pivot items for the row field
        PivotItemCollection items = pivotTable.RowFields[0].PivotItems;

        // Move the first item two positions down within the same parent node
        // Positive count moves down; isSameParent = true keeps the move inside the same parent
        items[0].Move(2, true);

        // Recalculate after moving items to reflect the new order
        pivotTable.CalculateData();

        // Save the workbook
        wb.Save("PivotItemMoveTwoPositions.xlsx");
    }
}
