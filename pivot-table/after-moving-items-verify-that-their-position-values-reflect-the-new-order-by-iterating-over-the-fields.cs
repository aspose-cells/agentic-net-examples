using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class VerifyPivotItemPositionsAfterMove
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
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
            sheet.Cells["B5"].PutValue("Potato");
            sheet.Cells["C5"].PutValue(70);

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add two fields to the row area: Category and SubCategory
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");

            // Add a data field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh and calculate to populate the pivot items
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            Console.WriteLine("=== Before Move ===");
            PrintPivotItemsPositions(pivotTable);

            // Move the first row field (Category) to position 1 (swap with SubCategory)
            // Using PivotFieldCollection.Move method as per the rule
            pivotTable.RowFields.Move(0, 1);

            // Refresh after moving fields
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            Console.WriteLine("\n=== After Move ===");
            PrintPivotItemsPositions(pivotTable);

            // Save the workbook (lifecycle rule)
            workbook.Save("PivotItemPositionAfterMove.xlsx");
        }

        // Helper method to iterate over row fields and their pivot items,
        // printing each item's Position property to verify order.
        private static void PrintPivotItemsPositions(PivotTable pivotTable)
        {
            for (int fieldIdx = 0; fieldIdx < pivotTable.RowFields.Count; fieldIdx++)
            {
                PivotField field = pivotTable.RowFields[fieldIdx];
                Console.WriteLine($"Row Field {fieldIdx}: {field.Name}");

                // Iterate over the PivotItems of the current field
                foreach (PivotItem item in field.PivotItems)
                {
                    // Position property reflects the item's index in the overall collection
                    Console.WriteLine($"    Item: {item.Name}, Position: {item.Position}");
                }
            }
        }
    }
}