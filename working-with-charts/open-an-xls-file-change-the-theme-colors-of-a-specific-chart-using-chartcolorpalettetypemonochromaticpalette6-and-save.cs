// Title: Changing a chart's theme to MonochromaticPalette6 in an existing XLS workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change chart color palette to MonochromaticPalette6 in an existing XLS file | Set Excel chart series colors programmatically with Aspose.Cells .NET | Load an XLS workbook, apply a monochrome palette to a chart, and save it using Aspose.Cells | How to apply a monochrome color palette to a chart in C# with Aspose.Cells | Update chart theme colors in an XLS workbook without opening Excel
// Tags: Aspose.Cells ChangeColors chart series | MonochromaticPalette6 chart color palette | modify chart theme XLS workbook | C# Aspose.Cells XlsSaveOptions | apply custom color palette to Excel chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartColorPaletteDemo
{
    // The example loads an existing XLS workbook, accesses the first worksheet and its first chart, changes the chart's series colors to the MonochromaticPalette6 palette via the ChangeColors method, and saves the workbook as a new XLS file using XlsSaveOptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths.
            // If command‑line arguments are supplied, use them; otherwise fall back to defaults.
            string inputPath = args.Length > 0 ? args[0] : "input.xls";
            string outputPath = args.Length > 1 ? args[1] : "output.xls";

            try
            {
                // Load the existing XLS workbook.
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook contains at least one worksheet.
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any worksheets.");
                    return;
                }

                // Use the first worksheet (you can modify this to target a specific sheet).
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart.
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine($"Worksheet \"{worksheet.Name}\" does not contain any charts.");
                    return;
                }

                // Select the first chart in the worksheet.
                // Replace the index with the appropriate chart number if needed.
                Chart chart = worksheet.Charts[0];

                // Change the color palette of the chart's series collection to MonochromaticPalette6.
                // The ChangeColors method is defined on SeriesCollection.
                chart.NSeries.ChangeColors(ChartColorPaletteType.MonochromaticPalette6);
                Console.WriteLine("Chart color palette changed to MonochromaticPalette6.");

                // Save the modified workbook as an XLS file.
                // XlsSaveOptions is used to ensure the correct format and to demonstrate explicit save options.
                XlsSaveOptions saveOptions = new XlsSaveOptions();
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
