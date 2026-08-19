// Title: Remove Trendlines from Excel Charts and Export as PNG with Aspose.Cells for .NET
// Description: Loads an XLS workbook, iterates through every worksheet and chart, deletes all trendlines from each series (using reflection for compatibility), saves the cleaned workbook, and exports each chart as a PNG image.
// Keywords: Aspose.Cells remove trendlines | Excel chart export PNG .NET | delete chart trendlines C# | Aspose.Cells chart image without trendlines | reflection trendline removal Aspose | batch process Excel charts | save modified workbook Aspose.Cells
// Common Searches: how to delete trendlines from all charts in an XLS file using Aspose.Cells | export Excel chart as PNG after removing trendlines .NET | remove chart trendlines programmatically Aspose.Cells | save workbook after chart modifications Aspose.Cells | C# code to export chart image without trendlines
// Developer Intent: Load an existing XLS file, strip every trendline from each chart, export the cleaned charts as PNG images, and write the updated workbook to a new file.
// Use Cases: Prepare presentation‑ready chart images by eliminating trendlines before export. | Automate cleanup of legacy Excel reports for bulk image generation. | Ensure compatibility with downstream systems that cannot interpret trendline objects.
// AI Prompts: Show a version of this code that removes trendlines without using reflection. | Suggest how to export charts at higher DPI while keeping trendlines removed. | Explain how to handle workbooks that contain no charts or no trendlines gracefully.

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Loads an XLS workbook, iterates through every worksheet and chart, deletes all trendlines from each series (using reflection for compatibility), saves the cleaned workbook, and exports each chart as a PNG image.
class RemoveTrendlinesAndExportImage
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.xls";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all charts on the current worksheet
                for (int chartIndex = 0; chartIndex < worksheet.Charts.Count; chartIndex++)
                {
                    Chart chart = worksheet.Charts[chartIndex];

                    // Remove every trendline from each series of the chart (using reflection for compatibility)
                    foreach (Series series in chart.NSeries)
                    {
                        PropertyInfo trendlinesProp = series.GetType().GetProperty("Trendlines");
                        if (trendlinesProp != null)
                        {
                            object trendlinesObj = trendlinesProp.GetValue(series);
                            if (trendlinesObj != null)
                            {
                                // TrendlineCollection implements IList, so we can treat it as such
                                var trendlines = trendlinesObj as System.Collections.IList;
                                if (trendlines != null && trendlines.Count > 0)
                                {
                                    // Remove trendlines in reverse order to avoid index shifting
                                    for (int t = trendlines.Count - 1; t >= 0; t--)
                                    {
                                        MethodInfo removeAt = trendlines.GetType().GetMethod("RemoveAt");
                                        removeAt?.Invoke(trendlines, new object[] { t });
                                    }
                                }
                            }
                        }
                    }

                    // Export the chart (now without trendlines) to an image file
                    try
                    {
                        string imageFileName = $"Chart_{worksheet.Name}_{chartIndex}.png";
                        chart.ToImage(imageFileName, ImageType.Png);
                        Console.WriteLine($"Chart exported to \"{imageFileName}\".");
                    }
                    catch (Exception imgEx)
                    {
                        Console.WriteLine($"Failed to export chart image: {imgEx.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
