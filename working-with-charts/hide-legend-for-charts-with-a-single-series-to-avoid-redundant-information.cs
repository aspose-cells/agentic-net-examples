// Title: Hide the legend of single‑series Excel charts with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide chart legend when only one series is present | remove legend from Excel chart programmatically using Aspose.Cells .NET | how to suppress redundant legend in a single‑series chart with Aspose.Cells | set ShowLegend false for charts with one series in an Aspose.Cells workbook | iterate through worksheets and charts to hide legend in .NET Excel automation
// Tags: Aspose.Cells ShowLegend property | single series chart legend suppression | Excel chart legend hide .NET | worksheet chart iteration Aspose.Cells | conditional legend visibility Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace Example
{
    // The example loads a workbook, iterates over each worksheet and its charts, checks for charts that contain exactly one data series, disables the legend by setting ShowLegend to false, and saves the updated file.
    class HideLegendForSingleSeriesChart
    {
        static void Main()
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through worksheets and their charts
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Hide legend if the chart contains only one series
                        if (chart.NSeries.Count == 1)
                        {
                            // Use ShowLegend property to control visibility
                            chart.ShowLegend = false;
                        }
                    }
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
