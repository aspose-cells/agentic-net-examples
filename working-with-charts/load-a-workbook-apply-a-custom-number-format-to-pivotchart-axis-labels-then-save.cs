using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace MyApp
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "PivotChartSample.xlsx";
                const string outputPath = "PivotChartFormatted.xlsx";

                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook containing the PivotChart
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Verify that the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Access the first chart on the worksheet (the PivotChart)
                Chart pivotChart = worksheet.Charts[0];

                // Apply custom number format to the category (X) axis tick labels
                pivotChart.CategoryAxis.TickLabels.NumberFormat = "$#,##0";

                // Apply custom number format to the value (Y) axis tick labels
                pivotChart.ValueAxis.TickLabels.NumberFormat = "0.00%";

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}