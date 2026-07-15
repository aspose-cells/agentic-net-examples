using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class DeletePivotTableDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot tables
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Apple");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["B4"].PutValue(150);

                // Add three pivot tables to the worksheet
                sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                sheet.PivotTables.Add("A1:B4", "D10", "PivotTable2");
                sheet.PivotTables.Add("A1:B4", "D20", "PivotTable3");

                // Remove the second pivot table (index 1) using RemoveAt
                sheet.PivotTables.RemoveAt(1);

                // Verify the remaining count
                Console.WriteLine("Remaining PivotTables count: " + sheet.PivotTables.Count);

                // Save the workbook
                string outputPath = "PivotTableRemoved.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            DeletePivotTableDemo.Run();
        }
    }
}