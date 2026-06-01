using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPivotChartRefresh
{
    public class RefreshPivotChartExample
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Paths for input and output workbooks
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "RefreshedWorkbook.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook containing a PivotTable and a PivotChart
                Workbook workbook = new Workbook(inputPath);

                // Assume the first worksheet holds the source data, pivot table and pivot chart
                Worksheet worksheet = workbook.Worksheets[0];

                // ----- Modify the source data that the pivot table uses -----
                // Example: change some values in the data range
                worksheet.Cells["B2"].PutValue(1500); // original value changed
                worksheet.Cells["B3"].PutValue(2500); // original value changed

                // ----- Refresh all pivot tables in the workbook -----
                // This updates the pivot cache with the new source data
                workbook.Worksheets.RefreshPivotTables();

                // ----- Refresh the pivot chart so it reflects the updated pivot table -----
                // Assuming the chart is the first chart in the worksheet
                if (worksheet.Charts.Count > 0)
                {
                    Chart pivotChart = worksheet.Charts[0];
                    // Refreshes chart's data from its associated pivot table
                    pivotChart.RefreshPivotData();
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook refreshed and saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}