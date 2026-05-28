using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class DeletePivotTableExample
    {
        public static void Run()
        {
            try
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

                // Remove the second pivot table (index 1)
                sheet.PivotTables.RemoveAt(1);

                // Verify the remaining count
                Console.WriteLine("Remaining Pivot Tables Count: " + sheet.PivotTables.Count);

                // Save the workbook
                string outputPath = "PivotTableRemoved.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DeletePivotTableExample.Run();
        }
    }
}