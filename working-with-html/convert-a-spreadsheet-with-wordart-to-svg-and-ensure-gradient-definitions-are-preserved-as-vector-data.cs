// Title: Convert an Excel workbook containing WordArt to SVG while preserving gradient fills using Aspose.Cells for .NET
// Common Searches: how to export an Excel sheet that contains WordArt to SVG using Aspose.Cells C# | preserving WordArt gradient fills when converting Excel to SVG with Aspose.Cells | Aspose.Cells ImageSaveOptions SaveFormat.Svg example for WordArt objects | export each worksheet as separate SVG files in C# with Aspose.Cells | C# code to save Excel workbook as SVG and keep vector gradients
// Tags: Aspose.Cells ImageSaveOptions SVG export | retain gradient fills during SVG conversion | C# export each worksheet as separate SVG | one-page-per-sheet SVG generation | save Excel workbook as vector SVG with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, checks its existence, configures ImageSaveOptions to export each worksheet as a separate SVG with one page per sheet, creates the output directory if needed, and saves the workbook to SVG while preserving WordArt gradient fills as vector data.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing WordArt objects
            Workbook workbook = new Workbook(inputPath);

            // Configure SVG export using ImageSaveOptions (the recommended API)
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Svg)
            {
                // Export each worksheet to a separate SVG file
                ImageOrPrintOptions = { OnePagePerSheet = true }
            };

            // Define output file path
            const string outputPath = "output.svg";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook to SVG
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved as SVG to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
