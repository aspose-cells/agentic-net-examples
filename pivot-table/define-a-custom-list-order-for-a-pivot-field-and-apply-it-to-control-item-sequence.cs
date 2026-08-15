// Title: Set a Custom Row Item Order in an Aspose.Cells Pivot Table (C#)
// Description: This C# example creates a workbook, adds fruit‑quantity data, builds a pivot table, places the "Fruit" field in the row area, and defines a specific item sequence (Banana, Apple, Orange, Pear) by setting each PivotItem's PositionInSameParentNode property. The pivot is refreshed and saved as PivotTable_CustomItemOrder.xlsx.
// Keywords: Aspose.Cells | C# | PivotTable | custom item order | PositionInSameParentNode | row field sequence | programmatic pivot sorting | Excel pivot custom sort | Aspose.Cells example | set pivot item position
// Common Searches: Aspose.Cells set custom order for pivot row items | C# change pivot table item sequence programmatically | PositionInSameParentNode usage Aspose.Cells | how to reorder pivot table rows in .NET | custom sort order for Excel pivot table using Aspose.Cells
// Developer Intent: Define and apply a specific ordering for the items of a pivot table row field using Aspose.Cells for .NET.
// Use Cases: Align product categories with marketing hierarchy in sales dashboards | Display fiscal months in custom order for financial reporting | Prioritize service types in a support ticket analysis pivot | Arrange regional sales zones based on strategic importance | Create a custom order for survey response categories in analytics
// AI Prompts: Write C# code with Aspose.Cells that reorders pivot row items based on a user‑provided list. | Explain the PositionInSameParentNode property and how to reset pivot sorting to default. | Provide a generic method to apply custom ordering to multiple pivot fields in a workbook using Aspose.Cells. | Show how to retrieve and modify PivotItem positions dynamically from a CSV list. | Demonstrate error handling when a specified item does not exist in the pivot field.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsCustomPivotOrder
{
    // This C# example creates a workbook, adds fruit‑quantity data, builds a pivot table, places the "Fruit" field in the row area, and defines a specific item sequence (Banana, Apple, Orange, Pear) by setting each PivotItem's PositionInSameParentNode property. The pivot is refreshed and saved as PivotTable_CustomItemOrder.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            // Column A: Fruit, Column B: Quantity
            sheet.Cells["A1"].Value = "Fruit";
            sheet.Cells["B1"].Value = "Quantity";

            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 10;

            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["B3"].Value = 20;

            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B4"].Value = 15;

            sheet.Cells["A5"].Value = "Pear";
            sheet.Cells["B5"].Value = 5;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "FruitPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Fruit" field to the row area and "Quantity" to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // ------------------------------------------------------------
            // Define a custom order for the "Fruit" row field.
            // Desired order: Banana, Apple, Orange, Pear
            // ------------------------------------------------------------
            // Retrieve the row field
            PivotField fruitField = pivotTable.RowFields["Fruit"];
            // Access its PivotItem collection
            PivotItemCollection items = fruitField.PivotItems;

            // Helper method to set the position of an item within the same parent node
            void SetPosition(string itemName, int position)
            {
                if (items[itemName] != null)
                {
                    items[itemName].PositionInSameParentNode = position;
                }
            }

            // Apply the custom order
            SetPosition("Banana", 0); // First
            SetPosition("Apple", 1);  // Second
            SetPosition("Orange", 2); // Third
            SetPosition("Pear", 3);   // Fourth

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_CustomItemOrder.xlsx");
        }
    }
}
