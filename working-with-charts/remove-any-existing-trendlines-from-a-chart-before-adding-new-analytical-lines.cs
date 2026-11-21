// Title: How to remove all existing trendlines from an Excel chart and add new linear trendlines using Aspose.Cells for .NET (C#)
// Common Searches: C# code to delete all trendlines from a chart with Aspose.Cells | Aspose.Cells remove existing chart trendlines before adding new ones | clear trendlines in an Excel chart using the Aspose.Cells .NET API | add a linear trendline after clearing previous trendlines in Aspose.Cells C# | use reflection to access Trendlines property of a chart series in Aspose.Cells
// Tags: remove chart trendlines Aspose.Cells | add linear trendline Aspose.Cells C# | trendline collection reflection Aspose.Cells | chart series trendline management .NET | excel workbook trendline cleanup Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTrendlineExample
{
    // The example loads an Excel workbook, checks for a worksheet and a chart, iterates through each series to clear any existing trendlines via reflection (handling both Trendlines and TrendlineCollection properties), then adds a linear trendline to each series, and finally saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure there is at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("Error: The workbook does not contain any worksheets.");
                    return;
                }

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts found in the first worksheet.");
                    return;
                }

                // Access the first chart
                Chart chart = worksheet.Charts[0];

                // -------------------------------------------------
                // Remove all existing trendlines from every series
                // -------------------------------------------------
                foreach (Series series in chart.NSeries)
                {
                    try
                    {
                        // Some Aspose.Cells versions expose Trendlines via TrendlineCollection property
                        // Use whichever property is available at runtime
                        var trendlines = GetTrendlineCollection(series);
                        trendlines?.Clear();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to clear trendlines for a series – {ex.Message}");
                    }
                }

                // -------------------------------------------------
                // Add new analytical trendlines (example: linear)
                // -------------------------------------------------
                foreach (Series series in chart.NSeries)
                {
                    try
                    {
                        var trendlines = GetTrendlineCollection(series);
                        trendlines?.Add(TrendlineType.Linear);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Unable to add trendline to a series – {ex.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static TrendlineCollection GetTrendlineCollection(Series series)
        {
            // Prefer the 'Trendlines' property if it exists
            var type = series.GetType();

            var trendlinesProp = type.GetProperty("Trendlines");
            if (trendlinesProp != null)
            {
                return trendlinesProp.GetValue(series) as TrendlineCollection;
            }

            // Fallback to 'TrendlineCollection' property for older versions
            var trendlineCollectionProp = type.GetProperty("TrendlineCollection");
            if (trendlineCollectionProp != null)
            {
                return trendlineCollectionProp.GetValue(series) as TrendlineCollection;
            }

            // If neither property exists, return null
            return null;
        }
    }
}
