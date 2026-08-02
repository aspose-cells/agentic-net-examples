using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRemovalDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(200);

            // Add a pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table (row and data fields)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Remove the pivot table using the Remove method
            sheet.PivotTables.Remove(pivotTable);

            // Verify that no pivot tables remain in the worksheet
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Pivot table removed successfully. No remaining pivot tables.");
            }
            else
            {
                Console.WriteLine("Pivot table removal failed. Remaining count: " + sheet.PivotTables.Count);
            }

            // Save the workbook to a file
            workbook.Save("PivotTableRemoved.xlsx");
        }
    }
}