using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class DisablePivotRefreshOnOpen
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1000);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(2000);
                sheet.Cells["A4"].PutValue("Orange");
                sheet.Cells["B4"].PutValue(3000);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table (rows and data fields)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales column

                // Disable automatic refresh when the workbook is opened
                pivotTable.RefreshDataOnOpeningFile = false;

                // Define output file path
                string outputPath = "PivotTable_NoAutoRefresh.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            DisablePivotRefreshOnOpen.Run();
        }
    }
}