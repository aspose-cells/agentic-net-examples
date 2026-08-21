// Title: Aspose.Cells C# – Move a PivotItem to a Different Parent Field in a Pivot Table
// Description: Creates a workbook with hierarchical data (Category → SubCategory), builds a pivot table, and uses PivotItem.Move with isSameParent = false to reassign the first sub‑category item to another parent. The example refreshes the pivot, recalculates, and saves the result.
// Keywords: Aspose.Cells | C# | PivotItem.Move | isSameParent false | pivot table hierarchy | change pivot item parent | row field reordering | programmatic pivot manipulation | .NET pivot table example
// Common Searches: Aspose.Cells move pivot item to another parent | PivotItem.Move isSameParent false C# example | reassign subcategory in Aspose.Cells pivot table | change parent field of a pivot item programmatically | how to relocate a pivot item in .NET
// Developer Intent: Reassign a specific PivotItem to a different parent row field in a pivot table using Aspose.Cells for .NET.
// Use Cases: Fix mis‑categorized rows without rebuilding the entire pivot. | Enable interactive UI actions that let users drag a sub‑category to a new category. | Adjust hierarchical reporting structures after data merges or splits at runtime.
// AI Prompts: Write C# code that moves a PivotItem to a new parent field with Aspose.Cells and updates the pivot. | Explain the effect of the isSameParent parameter in PivotItem.Move. | Show how to verify the new parent of a moved PivotItem by inspecting the RowFields collection.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemMoveDemo
{
    // Creates a workbook with hierarchical data (Category → SubCategory), builds a pivot table, and uses PivotItem.Move with isSameParent = false to reassign the first sub‑category item to another parent. The example refreshes the pivot, recalculates, and saves the result.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a hierarchy (Category -> SubCategory)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("SubCategory");
            sheet.Cells["C1"].PutValue("Amount");

            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["C2"].PutValue(120);

            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue("Banana");
            sheet.Cells["C3"].PutValue(80);

            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue("Carrot");
            sheet.Cells["C4"].PutValue(50);

            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue("Tomato");
            sheet.Cells["C5"].PutValue(70);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add both fields to the row area to create a parent‑child relationship
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");      // Parent field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory"); // Child field

            // Add the data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh and calculate the pivot table so that items are generated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Access the PivotItemCollection of the child field (SubCategory)
            PivotItemCollection subItems = pivotTable.RowFields["SubCategory"].PivotItems;

            // Example: Move the first sub‑category item ("Apple") to a different parent node
            // Using count = 0 (no up/down shift) and isSameParent = false to indicate a different parent
            subItems[0].Move(0, false);

            // Recalculate after the move operation
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotItemMovedToDifferentParent.xlsx");
        }
    }
}
