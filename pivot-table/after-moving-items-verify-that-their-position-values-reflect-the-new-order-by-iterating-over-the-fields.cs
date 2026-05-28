using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    class VerifyPivotFieldPositions
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
            sheet.Cells["B5"].PutValue("Broccoli");
            sheet.Cells["C5"].PutValue(70);

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add two fields to the row area: Category and SubCategory
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");

            // Add the Amount field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh and calculate the pivot table so that items are generated
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Display field order and their Position property before moving
            Console.WriteLine("Before Move:");
            for (int i = 0; i < pivotTable.RowFields.Count; i++)
            {
                PivotField field = pivotTable.RowFields[i];
                Console.WriteLine($"Field Index: {i}, Name: {field.Name}, Position: {field.Position}");
            }

            // Move the first row field (Category) to position 1 (second place)
            pivotTable.RowFields.Move(0, 1);

            // Refresh again to ensure internal state is consistent
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Display field order and their Position property after moving
            Console.WriteLine("\nAfter Move:");
            for (int i = 0; i < pivotTable.RowFields.Count; i++)
            {
                PivotField field = pivotTable.RowFields[i];
                Console.WriteLine($"Field Index: {i}, Name: {field.Name}, Position: {field.Position}");
            }

            // Verify that each PivotItem within the fields also reflects correct ordering
            Console.WriteLine("\nPivot Items Position Verification:");
            foreach (PivotField field in pivotTable.RowFields)
            {
                Console.WriteLine($"Field: {field.Name}");
                foreach (PivotItem item in field.PivotItems)
                {
                    // Position property of PivotItem reflects its order among all items
                    Console.WriteLine($"  Item: {item.Name}, Position: {item.Position}");
                }
            }

            // Save the workbook (output file name can be adjusted as needed)
            workbook.Save("PivotFieldPositionVerification.xlsx");
        }
    }
}