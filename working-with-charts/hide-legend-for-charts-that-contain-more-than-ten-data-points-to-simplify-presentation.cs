// Title: Hide chart legend in Aspose.Cells for .NET when a chart contains more than ten data points
// Common Searches: Aspose.Cells C# hide chart legend if series has more than 10 points | How to programmatically remove legend from Excel chart using Aspose.Cells when data points exceed ten | Conditional legend visibility in Aspose.Cells .NET based on total data point count | C# Aspose.Cells iterate charts and hide legend for large data sets | Set legend size to zero in Aspose.Cells chart when point count > 10
// Tags: chart legend hiding based on data point count Aspose.Cells | Aspose.Cells hide legend for charts with many points | set legend size zero Aspose.Cells C# | legend visibility control based on point count Aspose.Cells | iterate worksheet charts Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace HideChartLegends
{
    // The example loads an Excel workbook, walks through each worksheet and its charts, counts the total data points across all series, and hides the chart legend by setting its position to NotDocked and its size to zero when the count exceeds ten; otherwise it restores the legend to the default position with automatic sizing, then saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.xlsx";

            // Verify that the input file exists before attempting to load it.
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputFile);

                // Iterate through each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each chart on the worksheet.
                    foreach (Chart chart in sheet.Charts)
                    {
                        try
                        {
                            int totalDataPoints = 0;

                            // Sum the number of points across all series in the chart.
                            foreach (Series series in chart.NSeries)
                            {
                                // Some series may have a null Points collection; guard against it.
                                if (series.Points != null)
                                {
                                    totalDataPoints += series.Points.Count;
                                }
                            }

                            Console.WriteLine($"Worksheet \"{sheet.Name}\", Chart \"{chart.Name}\": {totalDataPoints} data points.");

                            // Hide the legend if the chart has more than ten data points.
                            if (totalDataPoints > 10)
                            {
                                // Hide legend by making it not docked and setting its size to zero.
                                chart.Legend.Position = LegendPositionType.NotDocked;
                                chart.Legend.IsAutomaticSize = false;
                                chart.Legend.Width = 0;
                                chart.Legend.Height = 0;
                                Console.WriteLine("  -> Legend hidden (more than 10 data points).");
                            }
                            else
                            {
                                // Ensure the legend is visible for charts with 10 or fewer points.
                                // Reset to default position and automatic sizing.
                                chart.Legend.Position = LegendPositionType.Right;
                                chart.Legend.IsAutomaticSize = true;
                            }
                        }
                        catch (Exception exChart)
                        {
                            Console.WriteLine($"Error processing chart \"{chart.Name}\": {exChart.Message}");
                        }
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Processing complete. Modified workbook saved as \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
