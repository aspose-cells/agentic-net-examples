// Title: Apply a custom chart color palette to every chart in an Excel workbook and export it as XLS with Aspose.Cells for .NET
// AI Prompts: Load a .xlsx file, define an ARGB color array, loop through each worksheet and each chart, set the foreground color of every series using the palette, then save the workbook as an Excel 97‑2003 (.xls) file. | Programmatically enforce a consistent chart appearance by assigning predefined colors to all series in all charts of a workbook, handling missing files, and converting the result to the legacy XLS format using Aspose.Cells.
// Common Searches: how to set the same color scheme for all charts in an Excel file using Aspose.Cells C# | Aspose.Cells change chart series colors programmatically across multiple worksheets | convert a modified workbook to XLS format after updating chart colors with Aspose.Cells | apply ARGB palette to chart series in .NET Aspose.Cells example
// Tags: apply custom chart series colors Aspose.Cells | set chart series foreground ARGB palette | save workbook as Excel 97-2003 XLS | iterate worksheets and charts programmatically | uniform chart color palette across workbook

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The code loads an existing XLSX workbook, creates a custom ARGB color palette, iterates through every worksheet and each chart, assigns palette colors to the foreground of each series, and finally saves the updated workbook in the legacy Excel 97‑2003 (.xls) format.
class UniformChartPalette
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define a uniform custom color palette (ARGB values)
            int[] customPalette = new int[]
            {
                Color.FromArgb(0, 112, 192).ToArgb(),   // Blue
                Color.FromArgb(255, 192, 0).ToArgb(),   // Orange
                Color.FromArgb(112, 173, 71).ToArgb(),  // Green
                Color.FromArgb(255, 0, 0).ToArgb(),     // Red
                Color.FromArgb(128, 0, 128).ToArgb()    // Purple
            };

            // Iterate through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Apply custom colors to each series in the chart
                        for (int i = 0; i < chart.NSeries.Count; i++)
                        {
                            Series series = chart.NSeries[i];
                            int colorArgb = customPalette[i % customPalette.Length];
                            // Set the foreground color of the series
                            series.Area.ForegroundColor = Color.FromArgb(colorArgb);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing chart on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Save the modified workbook as XLS (Excel 97-2003 format)
            string outputPath = "output.xls";
            workbook.Save(outputPath, SaveFormat.Excel97To2003);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
