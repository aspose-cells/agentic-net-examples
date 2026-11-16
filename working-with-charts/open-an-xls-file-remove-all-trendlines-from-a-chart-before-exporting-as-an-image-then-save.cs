// Title: How to remove all trendlines from charts in an XLS workbook and export each chart as a PNG image using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove trendlines from chart series in an existing XLS file | Export chart to PNG after deleting trendlines with Aspose.Cells .NET | Iterate through worksheets and charts to clear trendlines in a legacy XLS workbook | Save modified XLS workbook without chart trendlines using Aspose.Cells | C# code to extract charts as images after removing trendlines from Excel file
// Tags: remove chart trendlines Aspose.Cells .NET | export chart as PNG Aspose.Cells | iterate worksheets charts Aspose.Cells C# | clear series trendlines via reflection C# | save XLS workbook without trendlines Aspose.Cells

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an XLS workbook, loops through every worksheet and chart, clears any trendlines from each series (using reflection when necessary), saves each chart as a PNG image, and finally writes the updated workbook back to disk without the original trendlines.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.xls";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            int chartCounter = 0;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Remove all trendlines from each series in the chart (if supported)
                        foreach (Series series in chart.NSeries)
                        {
                            var trendlinesProp = series.GetType().GetProperty("Trendlines");
                            if (trendlinesProp != null)
                            {
                                var trendlines = trendlinesProp.GetValue(series) as IList;
                                trendlines?.Clear();
                            }
                        }

                        // Export the chart as a PNG image (default format is PNG)
                        string chartImagePath = $"chart_{chartCounter}.png";
                        chart.ToImage(chartImagePath);
                        chartCounter++;
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart in sheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook (trendlines removed if possible)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
