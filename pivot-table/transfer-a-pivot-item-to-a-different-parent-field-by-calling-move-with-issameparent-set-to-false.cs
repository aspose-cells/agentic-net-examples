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
            // Columns: Category, SubCategory, Amount
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
            sheet.Cells["C4"].PutValue(150);

            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue("Broccoli");
            sheet.Cells["C5"].PutValue(130);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add two row fields: Category and SubCategory
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");

            // Add a data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh and calculate to populate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ------------------------------------------------------------
            // Transfer a pivot item from the first row field (Category) to
            // the second row field (SubCategory) by using Move with
            // isSameParent = false.
            // ------------------------------------------------------------

            // Get the PivotItemCollection of the first row field (Category)
            PivotItemCollection categoryItems = pivotTable.RowFields["Category"].PivotItems;

            // Choose a specific item to move, e.g., the item "Fruit"
            PivotItem fruitItem = categoryItems["Fruit"];

            // Move the item to a different parent (SubCategory field)
            // count = 0 (no up/down movement within the new parent)
            // isSameParent = false indicates a different parent node.
            fruitItem.Move(0, false);

            // Recalculate after the move operation
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotItemMovedToDifferentParent.xlsx");
        }
    }
}