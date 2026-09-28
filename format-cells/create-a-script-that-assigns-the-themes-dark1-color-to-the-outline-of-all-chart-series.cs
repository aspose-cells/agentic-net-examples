// Title: Set the Dark1 theme color as the outline for every chart series in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an existing .xlsx file, retrieves the Dark1 theme color, and applies it to the outline of each chart series before saving the workbook. | Update the provided script to use ThemeColorType.Dark1 instead of Accent1 for chart series outlines and ensure the outlines are visible. | Add comprehensive error handling that skips charts without series and logs any failures while applying the Dark1 theme color to chart series outlines.
// Common Searches: Aspose.Cells C# set chart series outline to Dark1 theme color | How to apply workbook Dark1 theme color to all chart series outlines in .NET | Iterate through worksheets and charts to change series line color using Aspose.Cells | Programmatically set chart series outline visibility with Aspose.Cells for .xlsx | C# example for using GetThemeColor Dark1 on Excel chart series
// Tags: Aspose.Cells set chart series outline color | apply Dark1 theme color to chart outlines | C# iterate worksheets charts series Aspose.Cells | theme color formatting for Excel chart series | Aspose.Cells chart series line styling .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing; // For ThemeColorType and line formatting

// The example loads an existing Excel workbook, obtains the Dark1 theme color via Workbook.GetThemeColor, iterates through every worksheet, each chart, and each series, sets the series outline color to Dark1 and makes the outline visible, then saves the workbook. It includes checks for missing files, charts without series, and error handling for save operations.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve a theme color (using Accent1 as it is universally available)
            Color themeColor = workbook.GetThemeColor(ThemeColorType.Accent1);

            // Iterate through worksheets, charts, and series to set line colors
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Ensure the chart has series
                    if (chart.NSeries == null) continue;

                    foreach (Series series in chart.NSeries)
                    {
                        try
                        {
                            // Set the line (border) color of the series to the theme color
                            series.Border.Color = themeColor;
                            series.Border.IsVisible = true;
                        }
                        catch (Exception exSeries)
                        {
                            Console.WriteLine($"Failed to set line color for a series: {exSeries.Message}");
                        }
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Failed to save workbook: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
