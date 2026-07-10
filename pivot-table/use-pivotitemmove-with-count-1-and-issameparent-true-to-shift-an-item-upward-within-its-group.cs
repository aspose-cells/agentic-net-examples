using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemMoveDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(200);
            sheet.Cells["B5"].PutValue(250);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Refresh and calculate to populate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Access the pivot items of the "Category" row field
            PivotItemCollection items = pivotTable.RowFields["Category"].PivotItems;

            // Example: Move the second item ("Fruit") up by one position within its parent group
            // Count = -1 (move up), isSameParent = true (stay within the same parent node)
            if (items.Count > 1)
            {
                // Move the item at index 1 upward
                items[1].Move(-1, true);
            }

            // Recalculate after moving the item
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotItemMoveUpDemo.xlsx");
        }
    }
}