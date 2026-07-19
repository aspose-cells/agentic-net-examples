// Title: C# – Remove Chart Trendlines and Export as Image with Aspose.Cells
// Description: Loads an XLS workbook, clears every trendline from all chart series, exports a chart to PNG, and saves the workbook as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart trendline removal | export Excel chart to image Aspose.Cells | delete trendlines programmatically | save workbook without trendlines | chart image generation .NET
// Common Searches: how to delete trendlines from Excel charts using Aspose.Cells | export chart to PNG after removing trendlines C# | remove all chart trendlines and save workbook Aspose.Cells | Aspose.Cells chart image export without trendlines
// Developer Intent: Strip every trendline from all charts in a workbook, create an image of a chart, and persist the cleaned file.
// Use Cases: Prepare presentation‑ready chart images by eliminating trendlines before export. | Distribute Excel files that contain only essential data visualizations, omitting trendlines. | Automate batch processing to clean charts and generate PNG snapshots for documentation.
// AI Prompts: Generate C# code with Aspose.Cells that removes trendlines from all chart series without using reflection. | Show how to export the first chart of each worksheet to JPEG after clearing trendlines, then save the workbook as .xlsb. | Explain safe detection of the Trendlines property and its removal in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTrendlineRemoval
{
    // Loads an XLS workbook, clears every trendline from all chart series, exports a chart to PNG, and saves the workbook as XLSX using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source XLS file
                string sourceFile = "input.xls";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file not found: {sourceFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourceFile);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    for (int chartIdx = 0; chartIdx < sheet.Charts.Count; chartIdx++)
                    {
                        Chart chart = sheet.Charts[chartIdx];

                        // Remove all trendlines from each series in the chart (if supported)
                        foreach (Series series in chart.NSeries)
                        {
                            // The Trendlines collection may not be available in older versions of Aspose.Cells.
                            // Guard against missing members using reflection.
                            var trendlinesProp = series.GetType().GetProperty("Trendlines");
                            if (trendlinesProp != null)
                            {
                                var trendlines = trendlinesProp.GetValue(series) as TrendlineCollection;
                                if (trendlines != null)
                                {
                                    while (trendlines.Count > 0)
                                    {
                                        trendlines.RemoveAt(0);
                                    }
                                }
                            }
                        }

                        // Export the first chart (or any chart) to an image file
                        // The file extension determines the image format (e.g., .png)
                        if (chartIdx == 0) // export only once per worksheet for demonstration
                        {
                            try
                            {
                                string imagePath = $"Chart_{sheet.Name}_{chartIdx}.png";
                                chart.ToImage(imagePath);
                                Console.WriteLine($"Chart exported to image: {imagePath}");
                            }
                            catch (Exception imgEx)
                            {
                                Console.WriteLine($"Failed to export chart image: {imgEx.Message}");
                            }
                        }
                    }
                }

                // Save the modified workbook
                string outputFile = "output.xlsx";
                workbook.Save(outputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved without trendlines: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
