// Title: Render Excel worksheets with WordArt to high‑resolution PNG images while preserving gradient fills using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, verifies its existence, and uses Aspose.Cells to render each worksheet containing WordArt to a 300 dpi PNG with transparent background and gradient‑fill preservation. | Show how to configure ImageOrPrintOptions in Aspose.Cells to adjust DPI, enable transparency, and switch the output format from PNG to JPEG when rendering Excel sheets. | Write a snippet that iterates over all worksheets in a workbook, saves each as a separate PNG named after the sheet, and logs the generated file paths.
// Common Searches: how to export Excel sheet with WordArt to PNG preserving gradient colors using Aspose.Cells .NET | C# Aspose.Cells render worksheet to high DPI PNG with transparent background | set image resolution when converting Excel to PNG with Aspose.Cells | preserve WordArt gradient fill when converting Excel to image in .NET
// Tags: Aspose.Cells SheetRender to PNG with gradient fill | high DPI image export Aspose.Cells | transparent background PNG from Excel WordArt | ImageOrPrintOptions DPI setting Aspose.Cells | batch render worksheets to PNG Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, checks the input file, and uses Aspose.Cells SheetRender with high‑resolution ImageOrPrintOptions to export each worksheet containing WordArt to a 300 dpi PNG that retains gradient fills and transparency.
class WordArtToPngConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the Excel workbook that contains WordArt.
            Workbook workbook = new Workbook(inputPath);

            // Set up image rendering options.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Render the whole sheet on a single page.
                OnePagePerSheet = true,
                // High resolution to preserve gradient fill quality.
                HorizontalResolution = 300,
                VerticalResolution = 300,
                // Maximum quality.
                Quality = 100,
                // Preserve transparency if WordArt contains it.
                Transparent = true
                // Note: ImageFormat defaults to PNG, so no explicit setting is required.
            };

            // Loop through each worksheet and render it to a PNG file.
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                // Create a SheetRender object for the current sheet.
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);
                // Render the first (and only) page to an image.
                // The file name includes the sheet name for clarity.
                string outputPath = $"Sheet_{sheet.Name}.png";
                sheetRender.ToImage(0, outputPath);
                Console.WriteLine($"Rendered sheet \"{sheet.Name}\" to \"{outputPath}\".");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
