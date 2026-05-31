using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDeleteDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for pivot tables
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(150);

            // Add two pivot tables to demonstrate removal
            // First pivot table at D1
            sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            // Second pivot table at D10
            sheet.PivotTables.Add("A1:B4", "D10", "PivotTable2");

            // Verify count before removal
            Console.WriteLine("Pivot tables before removal: " + sheet.PivotTables.Count);

            // Delete the pivot table at zero‑based index 0 (the first pivot table)
            sheet.PivotTables.RemoveAt(0);

            // Verify count after removal
            Console.WriteLine("Pivot tables after removal: " + sheet.PivotTables.Count);

            // Save the workbook to a file
            workbook.Save("PivotTableRemoved.xlsx");
        }
    }
}