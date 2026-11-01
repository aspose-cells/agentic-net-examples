// Title: Change the color palette of a chart to MonochromePalette6 in an existing XLS workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing XLS file, retrieve the first chart, assign ChartColorPaletteType.MonochromePalette6 to its Palette property, and save the workbook with Aspose.Cells in C#. | Use reflection to set the Palette property of a chart to the MonochromePalette6 enum when the direct property is unavailable in the current Aspose.Cells version. | Programmatically update the theme colors of a chart in a legacy XLS workbook and write the modified file to a new location using Aspose.Cells for .NET.
// Common Searches: how to set chart palette to MonochromePalette6 with Aspose.Cells C# | change chart theme colors in an existing .xls file using Aspose.Cells .NET | using reflection to modify chart properties in Aspose.Cells for legacy Excel files
// Tags: Aspose.Cells chart palette assignment | MonochromePalette6 chart color theme .NET | update XLS chart theme using Aspose.Cells | reflection based chart property modification | change Excel chart colors programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLS workbook, accesses the first worksheet and its first chart, uses reflection to assign the ChartColorPaletteType.MonochromePalette6 enum to the chart's Palette property, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.xls";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing XLS file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count > 0)
            {
                // Get the first chart
                Chart chart = worksheet.Charts[0];

                // Attempt to set the chart's palette using reflection (covers different library versions)
                try
                {
                    var paletteProp = chart.GetType().GetProperty("Palette");
                    if (paletteProp != null && paletteProp.CanWrite)
                    {
                        // Try to obtain the ChartColorPaletteType enum type via reflection
                        var enumType = chart.GetType().Assembly.GetType("Aspose.Cells.Charts.ChartColorPaletteType");
                        if (enumType != null)
                        {
                            var enumValue = Enum.Parse(enumType, "MonochromePalette6");
                            paletteProp.SetValue(chart, enumValue);
                            Console.WriteLine("Chart palette set to MonochromePalette6.");
                        }
                        else
                        {
                            Console.WriteLine("ChartColorPaletteType enum not found; cannot set palette.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Chart palette property not available in this Aspose.Cells version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set chart palette: {ex.Message}");
                }
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
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
