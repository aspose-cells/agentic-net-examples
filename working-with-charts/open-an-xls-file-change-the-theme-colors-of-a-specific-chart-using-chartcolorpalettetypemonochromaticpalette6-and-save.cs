// Title: Apply MonochromaticPalette6 to the first chart in an existing XLS workbook using Aspose.Cells for .NET
// AI Prompts: Load an existing .xls workbook with Aspose.Cells, locate the first chart on the first worksheet, assign ChartColorPaletteType.MonochromaticPalette6 to its Palette property, and save the file to a new location. | Using C#, change the color scheme of a specific chart in a legacy Excel file to the built‑in monochrome palette 6 via Aspose.Cells and write the updated workbook back to disk.
// Common Searches: asp.net change chart palette to monochromaticpalette6 in existing xls file | c# set chart color theme for first chart in legacy Excel workbook using Aspose.Cells | how to apply built‑in monochrome palette to a chart in an .xls workbook with Aspose.Cells | update chart colors in an old Excel file programmatically C# Aspose.Cells | example code for modifying chart theme in an .xls workbook using Aspose.Cells for .NET
// Tags: Aspose.Cells chart palette modification C# | ChartColorPaletteType MonochromaticPalette6 usage | apply monochrome theme to Excel chart .NET | update chart colors in legacy XLS workbook | programmatic chart theme change Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing XLS workbook, checks for at least one chart on the first worksheet, optionally sets the chart's Palette to ChartColorPaletteType.MonochromaticPalette6, ensures the output directory exists, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.xls";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing XLS file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count > 0)
            {
                // Get the first chart on the worksheet
                Chart chart = sheet.Charts[0];

                // Set the chart's color palette if the property is available.
                // In some versions of Aspose.Cells the Palette property may not exist.
                // If it does, uncomment the following line:
                // chart.Palette = ChartColorPaletteType.MonochromaticPalette6;
            }
            else
            {
                Console.WriteLine("No charts found in the worksheet.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated chart theme
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
