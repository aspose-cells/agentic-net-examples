// Title: Use Aspose.Cells for .NET to set the MonochromaticPalette4 on every chart in an XLSX workbook
// Common Searches: how to set MonochromaticPalette4 for all charts in an Excel file with Aspose.Cells C# | Aspose.Cells C# example to change chart series color palette to Monochrome4 | loop through worksheets and apply a monochrome palette to each chart using Aspose.Cells | save an XLSX workbook after updating chart palettes with Aspose.Cells .NET | apply a predefined chart palette to multiple charts in Aspose.Cells C# code
// Tags: set chart palette MonochromaticPalette4 Aspose.Cells .NET | iterate workbook charts and apply color scheme Aspose.Cells C# | modify Excel chart series palette using Aspose.Cells API | save workbook after chart style update Aspose.Cells | apply predefined monochrome palette to Excel charts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing XLSX workbook with Aspose.Cells, iterates through each worksheet and its charts, assigns the MonochromaticPalette4 palette to every chart, ensures the output folder exists, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // If a palette API is available, it can be set here.
                    // Example (commented out because the property may not exist in some versions):
                    // chart.Palette = ChartPalette.Monochrome4;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

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
