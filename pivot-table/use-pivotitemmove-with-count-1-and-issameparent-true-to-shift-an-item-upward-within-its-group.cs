// Title: Aspose.Cells for .NET – Move a Pivot Row Item Up Using PivotItem.Move(-1, true)
// Description: Creates a workbook, adds sample data, builds a pivot table, places the "Category" field in the row area, then moves the second row item ("Beta") one position upward within its parent group by calling PivotItem.Move(-1, true) and saves the result.
// Keywords: Aspose.Cells | PivotItem.Move | C# | move pivot row item up | reorder pivot items | same parent group | negative count | pivot table automation | Aspose.Cells example
// Common Searches: Aspose.Cells move pivot row item up | PivotItem.Move usage C# | How to reorder pivot items in Aspose.Cells | Shift pivot table row item upward | PivotItem.Move(-1, true) sample code
// Developer Intent: Programmatically shift a specific pivot row entry one position higher while preserving its parent hierarchy.
// Use Cases: Reorder category entries in a generated report to highlight priority items before exporting. | Adjust the display order of pivot items after data refresh to match business‑specific sequencing. | Customize pivot table layouts by moving items within their groups to comply with reporting standards.
// AI Prompts: Show how to move the third pivot item down two positions with PivotItem.Move in C#. | Provide code that moves a pivot column item to the top of its group using Aspose.Cells. | Explain how to loop through multiple pivot items and reorder them within the same parent using PivotItem.Move.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemMoveDemo
{
    // Creates a workbook, adds sample data, builds a pivot table, places the "Category" field in the row area, then moves the second row item ("Beta") one position upward within its parent group by calling PivotItem.Move(-1, true) and saves the result.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["A4"].PutValue("Gamma");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Refresh and calculate to populate the pivot items
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Get the collection of pivot items for the row field
            PivotItemCollection items = pivotTable.RowFields[0].PivotItems;

            // Example: Move the second item ("Beta") up by one position within its parent group
            // Count = -1 (move up), isSameParent = true (stay within the same parent node)
            items[1].Move(-1, true);

            // Recalculate after the move operation
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotItemMoveUpDemo.xlsx");
        }
    }
}
