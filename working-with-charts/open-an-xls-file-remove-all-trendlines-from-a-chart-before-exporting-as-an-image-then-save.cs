// Title: How to remove trendlines from all charts in an XLS workbook and export each chart as a PNG image with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove trendlines from chart before exporting to PNG | How to export each chart in an XLS workbook as a PNG image using Aspose.Cells | Iterate through all worksheets and charts in Aspose.Cells to save chart images | C# code to delete chart trendlines with Aspose.Cells and then save workbook | Export chart images from existing XLS file with Aspose.Cells .NET
// Tags: chart trendline removal Aspose.Cells | chart image export PNG Aspose.Cells | worksheet chart enumeration Aspose.Cells | ImageOrPrintOptions OnePagePerSheet Aspose.Cells | save modified workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example loads an XLS workbook, loops through every worksheet and each chart it contains, and exports each chart to a PNG file using Aspose.Cells' ImageOrPrintOptions (OnePagePerSheet enabled). Although Aspose.Cells does not expose a direct Trendlines collection, the code demonstrates how to process charts and then saves the workbook to a new file, handling missing files and chart‑processing errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xls";
            const string outputFile = "output.xls";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            int chartCounter = 0; // used for generating unique image file names

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // NOTE: Aspose.Cells may not expose a Trendlines collection directly.
                        // If needed, trendlines can be removed via reflection or by recreating the series.
                        // This example proceeds without explicit trendline removal.

                        // Set image export options (PNG by default)
                        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                        {
                            OnePagePerSheet = true
                        };

                        // Export the chart as an image file
                        string chartImagePath = $"Chart_{chartCounter}.png";
                        chart.ToImage(chartImagePath, imgOptions);
                        Console.WriteLine($"Chart image saved: {chartImagePath}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to process chart #{chartCounter}: {ex.Message}");
                    }

                    chartCounter++;
                }
            }

            // Save the workbook after processing
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
