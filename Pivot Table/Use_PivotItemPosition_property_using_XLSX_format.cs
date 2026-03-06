using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotItemPositionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["B3"].PutValue(2000);
            sheet.Cells["B4"].PutValue(3000);

            // Add a pivot table using the data range A1:B4, place it at E3, and name it "PivotTable1"
            int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[ptIndex];

            // Add the row field (Product) and the data field (Sales) to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Access the row field's PivotItems collection
            PivotField rowField = pivotTable.RowFields[0];
            var itemsSnapshot = rowField.PivotItems.Cast<PivotItem>().ToList();

            foreach (PivotItem item in itemsSnapshot)
            {
                // Display the original Position value
                Console.WriteLine($"Before - Name: {item.Name}, Position: {item.Position}");

                // Set Position to 0 (moves the item to the top of the global list)
                item.Position = 0;

                // Display the updated Position value
                Console.WriteLine($"After  - Name: {item.Name}, Position: {item.Position}");
            }

            // Refresh the pivot table data and recalculate after modifying positions
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook in XLSX format
            workbook.Save("PivotItemPositionDemo.xlsx");
        }
    }
}