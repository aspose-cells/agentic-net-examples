using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class RefreshAndCalculatePivotTable
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Access first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Process first pivot table if present
            if (worksheet.PivotTables.Count > 0)
            {
                PivotTable pivotTable = worksheet.PivotTables[0];
                pivotTable.RefreshData();   // Refresh data from source
                pivotTable.CalculateData(); // Recalculate pivot values
            }
            else
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}