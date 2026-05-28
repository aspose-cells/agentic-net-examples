using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRemovalDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["B4"].PutValue(80);
            sheet.Cells["A5"].PutValue("Banana");
            sheet.Cells["B5"].PutValue(70);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D2", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // Verify that a pivot table exists before removal
            Console.WriteLine("Pivot tables count before removal: " + sheet.PivotTables.Count);

            // Remove the pivot table using the Remove method
            sheet.PivotTables.Remove(pivot);

            // Verify that no pivot tables remain in the worksheet
            int remainingCount = sheet.PivotTables.Count;
            Console.WriteLine("Pivot tables count after removal: " + remainingCount);
            Console.WriteLine(remainingCount == 0
                ? "Verification passed: No pivot tables remain."
                : "Verification failed: Some pivot tables are still present.");

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            workbook.Save("PivotTableRemovedDemo.xlsx");
        }
    }
}