using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDeleteExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for pivot tables
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            // Add three pivot tables to the worksheet
            sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            sheet.PivotTables.Add("A1:B4", "D10", "PivotTable2");
            sheet.PivotTables.Add("A1:B4", "D20", "PivotTable3");

            // Remove the pivot table at zero‑based index 1 (the second pivot table)
            sheet.PivotTables.RemoveAt(1);

            // Optional: display remaining count to verify removal
            Console.WriteLine("Remaining Pivot Tables Count: " + sheet.PivotTables.Count);

            // Save the workbook
            workbook.Save("PivotTableRemoved.xlsx");
        }
    }
}