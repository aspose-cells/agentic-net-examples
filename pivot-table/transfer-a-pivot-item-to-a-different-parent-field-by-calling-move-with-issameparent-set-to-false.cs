using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemMoveDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["A5"].PutValue("Vegetable");

            sheet.Cells["B1"].PutValue("SubCategory");
            sheet.Cells["B2"].PutValue("Apple");
            sheet.Cells["B3"].PutValue("Banana");
            sheet.Cells["B4"].PutValue("Carrot");
            sheet.Cells["B5"].PutValue("Potato");

            sheet.Cells["C1"].PutValue("Sales");
            sheet.Cells["C2"].PutValue(120);
            sheet.Cells["C3"].PutValue(150);
            sheet.Cells["C4"].PutValue(80);
            sheet.Cells["C5"].PutValue(90);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add two row fields – they will become parent nodes for pivot items
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");      // Parent field 0
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");   // Parent field 1

            // Refresh data so that pivot items are generated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Get the first pivot item from the first parent field ("Category")
            PivotItemCollection categoryItems = pivotTable.RowFields["Category"].PivotItems;
            PivotItem firstCategoryItem = categoryItems[0]; // e.g., "Fruit"

            // Move the item to a different parent field ("SubCategory")
            // count = 0 (position unchanged within the new parent), isSameParent = false
            firstCategoryItem.Move(0, false);

            // Recalculate after the move operation
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotItemMovedToDifferentParent.xlsx");
        }
    }
}