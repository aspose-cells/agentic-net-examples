using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the worksheet that should contain the pivot table
                Worksheet worksheet = workbook.Worksheets["Sheet1"] ?? workbook.Worksheets[0];

                // Update source data (example range A2:B5)
                worksheet.Cells["B2"].PutValue(1500);
                worksheet.Cells["B3"].PutValue(2500);
                worksheet.Cells["B4"].PutValue(1800);
                worksheet.Cells["B5"].PutValue(2200);

                // Ensure there is at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Refresh the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}