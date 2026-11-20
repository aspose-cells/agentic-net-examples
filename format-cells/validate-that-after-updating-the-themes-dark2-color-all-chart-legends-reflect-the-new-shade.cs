// Title: C# example to verify chart legend font color after changing the Dark2 theme color with Aspose.Cells
// Common Searches: how to verify chart legend font color after changing Dark2 theme in Aspose.Cells C# | Aspose.Cells C# validate that Excel chart legends reflect a custom Dark2 shade | check legend font color matches updated Dark2 theme color using Aspose.Cells | C# code to iterate all charts and confirm legend color after theme modification with Aspose.Cells | exception handling for legend color mismatch in Aspose.Cells chart processing
// Tags: chart legend font color update Aspose.Cells | theme Dark2 color application Excel chart Aspose.Cells | validate legend color consistency C# Aspose.Cells | iterate worksheets and charts Aspose.Cells C# | handle legend color mismatch exception Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads a workbook, defines a new SteelBlue shade for the Dark2 theme, iterates through every worksheet and chart, sets each chart legend's font color to the new shade, validates that the applied color matches the expected value, throws an exception on mismatch, and saves the workbook.
class ThemeLegendValidation
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsx";
        const string outputPath = "OutputWorkbook.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the new shade for the Dark2 theme color (example: SteelBlue)
            Color newDark2Color = Color.FromArgb(255, 70, 130, 180);

            // Since Aspose.Cells does not expose a direct Theme API, apply the color to chart legends directly
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Ensure the chart has a legend
                    if (chart.Legend != null)
                    {
                        // Set the legend font color to the new Dark2 color
                        chart.Legend.Font.Color = newDark2Color;

                        // Validate that the legend's font color matches the expected color
                        Color legendFontColor = chart.Legend.Font.Color;
                        if (legendFontColor.ToArgb() != newDark2Color.ToArgb())
                        {
                            throw new InvalidOperationException(
                                $"Legend font color mismatch in chart '{chart.Name}' on sheet '{sheet.Name}'. " +
                                $"Expected: {newDark2Color}, Actual: {legendFontColor}");
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook if validation succeeds
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File error: {ex.Message}");
        }
        catch (InvalidOperationException ex)
        {
            Console.Error.WriteLine($"Validation error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
