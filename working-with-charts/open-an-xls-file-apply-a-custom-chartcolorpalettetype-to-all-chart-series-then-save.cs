// Title: Apply a custom ChartColorPaletteType to every chart series in an XLS workbook using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells change colors of all chart series in an existing XLS file | how to set ChartColorPaletteType for each chart in a workbook with Aspose.Cells | apply a monochromatic palette to charts in an XLS workbook programmatically | using Aspose.Cells C# to recolor chart series and save as XLS | example of Chart.NSeries.ChangeColors with XlsSaveOptions in .NET
// Tags: chart series color palette Aspose.Cells C# | ChartColorPaletteType ChangeColors usage | modify chart colors in XLS workbook | XlsSaveOptions preserve original format | iterate worksheets and charts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ApplyChartColorPalette
{
    // The sample validates command‑line arguments, loads an XLS workbook with Aspose.Cells, loops through each worksheet and its charts, applies a specified ChartColorPaletteType to all chart series via the ChangeColors method, and saves the updated workbook back to XLS using XlsSaveOptions to retain the original file format.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: ApplyChartColorPalette <input.xls> <output.xls> [Palette]");
                Console.WriteLine("Palette values: MonochromaticPalette1 .. MonochromaticPalette13");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];
            string paletteName = args.Length >= 3 ? args[2] : "MonochromaticPalette6";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            // Parse the palette enum value safely
            if (!Enum.TryParse(typeof(ChartColorPaletteType), paletteName, out object paletteObj))
            {
                Console.WriteLine($"Error: Invalid palette name \"{paletteName}\".");
                Console.WriteLine("Valid values are:");
                foreach (var name in Enum.GetNames(typeof(ChartColorPaletteType)))
                {
                    Console.WriteLine($"  {name}");
                }
                return;
            }

            ChartColorPaletteType palette = (ChartColorPaletteType)paletteObj;

            try
            {
                // Load the workbook (XLS format)
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Apply the selected color palette to the series collection
                        chart.NSeries.ChangeColors(palette);
                        Console.WriteLine($"Applied palette {palette} to chart \"{chart.Name}\" on sheet \"{sheet.Name}\".");
                    }
                }

                // Save using XlsSaveOptions to preserve XLS format
                XlsSaveOptions saveOptions = new XlsSaveOptions();
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
