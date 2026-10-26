// Title: Determine whether a chart point is on the secondary axis in an Excel chart using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if chart point is on secondary axis | how to use ChartPoint.IsInSecondaryPlot property in .NET | iterate through chart series points to find secondary plot points with Aspose.Cells | read secondary plot flag for Excel chart points using Aspose.Cells API | C# code example for detecting secondary axis points in Excel charts
// Tags: ChartPoint.IsInSecondaryPlot property Aspose.Cells | secondary axis detection in Excel chart .NET | iterate chart series points Aspose.Cells C# | Excel chart secondary plot identification | Aspose.Cells chart point analysis

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the first chart, iterates over each series and its points, and uses the ChartPoint.IsInSecondaryPlot property to determine if a point belongs to the secondary plot, outputting the result before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Check for charts on the worksheet
                if (worksheet.Charts.Count > 0)
                {
                    // Get the first chart (adjust index if needed)
                    Chart chart = worksheet.Charts[0];

                    // Iterate through all series in the chart
                    for (int s = 0; s < chart.NSeries.Count; s++)
                    {
                        var series = chart.NSeries[s];

                        // Iterate through all points in the current series
                        for (int p = 0; p < series.Points.Count; p++)
                        {
                            ChartPoint point = series.Points[p];

                            // Determine if the point belongs to the secondary plot
                            bool belongsToSecondary = point.IsInSecondaryPlot;

                            // Output the result
                            Console.WriteLine(
                                $"Series Index: {s}, Point Index: {p}, IsInSecondaryPlot: {belongsToSecondary}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No charts found on the worksheet.");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to the desired output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
