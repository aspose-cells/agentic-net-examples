// Title: Disable text wrapping for all chart data labels across every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through each worksheet and chart, enables data labels, and sets DataLabels.IsTextWrapped = false for every series. | Update an existing Excel file so that all chart series show data labels without text wrap, then save the modified workbook to a new location using Aspose.Cells. | Create a resilient C# routine that verifies the input file, processes all charts in all sheets to turn off label wrapping, and handles any errors during the operation.
// Common Searches: aspnet aspocells disable chart data label wrap for all worksheets | c# loop through all charts in a workbook and turn off data label text wrapping | batch update Excel chart series data labels to not wrap text using Aspose.Cells | how to set DataLabels.IsTextWrapped false for every chart in an Excel file with Aspose.Cells | process multiple worksheets to modify chart label wrapping in C#
// Tags: disable chart data label wrapping Aspose.Cells | iterate all worksheets and charts C# | set DataLabels.IsTextWrapped false | enable data labels for chart series Aspose.Cells | save workbook after chart label changes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, checks that the file exists, then iterates through every worksheet, each chart, and each series to ensure data labels are displayed and disables text wrapping for those labels before saving the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Iterate through all series of the chart
                    foreach (Series series in chart.NSeries)
                    {
                        try
                        {
                            // Enable data labels for the series if not already enabled
                            if (!series.DataLabels.ShowValue)
                                series.DataLabels.ShowValue = true;

                            // Disable text wrap for all data labels in the series
                            series.DataLabels.IsTextWrapped = false;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Warning: Unable to modify series in chart '{chart.Name}'. {ex.Message}");
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
