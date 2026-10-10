// Title: Hide a chart legend entry when its series color matches a specific RGB threshold using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates through every chart in a workbook and makes the corresponding legend entry invisible if the series fill color equals a supplied RGB value. | Update an existing Excel file with Aspose.Cells so that any chart series whose area foreground color matches a defined color has its legend entry hidden by setting the legend font to transparent.
// Common Searches: asp.net aspose.cells hide legend entry based on series color | c# change chart legend visibility when series fill matches a color | how to make a legend entry transparent in an Excel chart using Aspose.Cells | conditional formatting of chart legend entries by series color in .NET
// Tags: set legend entry font transparent Aspose.Cells | chart series color threshold handling .NET | conditional legend visibility Excel chart | modify chart legend entries programmatically | compare series area foreground color Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example loads or creates a workbook, defines an RGB color threshold, scans each worksheet and chart, checks each series' area foreground color, and if it matches the threshold, sets the associated legend entry's font color to transparent before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load workbook if file exists; otherwise create a new workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Define the color threshold (example: pure red)
            Color thresholdColor = Color.FromArgb(255, 0, 0);

            // Iterate through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Ensure the chart has series and a legend
                    if (chart.NSeries == null || chart.Legend == null)
                        continue;

                    int seriesCount = chart.NSeries.Count;
                    for (int i = 0; i < seriesCount; i++)
                    {
                        Series series = chart.NSeries[i];

                        // Retrieve the series color (using the series area foreground color)
                        Color seriesColor = series.Area.ForegroundColor;

                        // If the series color matches the threshold, modify the legend entry
                        if (seriesColor.ToArgb() == thresholdColor.ToArgb())
                        {
                            // Ensure the legend entry exists
                            if (i < chart.Legend.LegendEntries.Count)
                            {
                                LegendEntry legendEntry = chart.Legend.LegendEntries[i];

                                // Hide the legend entry by making its font transparent
                                legendEntry.Font.Color = Color.Transparent;
                            }
                        }
                    }
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
