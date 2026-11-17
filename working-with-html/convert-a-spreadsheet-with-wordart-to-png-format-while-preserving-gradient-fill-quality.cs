// Title: Export an Excel worksheet with WordArt to a high‑resolution PNG while preserving gradient fills using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export worksheet containing WordArt to PNG with gradient fill intact | how to render Excel WordArt as PNG without losing gradient quality in .NET | Aspose.Cells ImageOrPrintOptions settings for high quality PNG of shapes | convert first sheet of Excel file to PNG preserving WordArt gradients using C# | save Excel sheet as PNG image with WordArt gradient using Aspose.Cells
// Tags: Aspose.Cells WordArt PNG export | ImageOrPrintOptions gradient quality | SheetRender high resolution PNG | C# Excel to PNG shape preservation | OnePagePerSheet PNG rendering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, configures ImageOrPrintOptions (OnePagePerSheet, Quality = 100, Transparent = false) to keep gradient fills sharp, and uses SheetRender to convert the first worksheet—including WordArt objects—into a high‑resolution PNG file, creating the output folder if needed and handling missing files or runtime errors.
class WordArtToPng
{
    static void Main()
    {
        // Path to the source workbook containing WordArt
        string inputFile = "input.xlsx";

        // Path for the resulting PNG image
        string outputFile = "output.png";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputFile);

            // Configure image rendering options to preserve high‑quality gradient fills
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Render each sheet on a single page (helps keep shapes intact)
                OnePagePerSheet = true,

                // Ensure maximum rendering quality
                Quality = 100,

                // Disable transparency to avoid artifacts in gradient rendering
                Transparent = false
                // ImageFormat defaults to PNG, so explicit setting is unnecessary
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Render the first worksheet (index 0) to an image
            SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], imgOptions);
            sheetRender.ToImage(0, outputFile);

            Console.WriteLine($"Image saved to {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
