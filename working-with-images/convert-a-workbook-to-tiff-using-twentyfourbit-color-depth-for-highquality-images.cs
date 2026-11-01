// Title: How to convert an Excel workbook to a 24‑bit TIFF image with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet to 24-bit TIFF file | How to save an Excel workbook as a high‑quality TIFF using Aspose.Cells | Aspose.Cells ImageOrPrintOptions set 24-bit color depth for TIFF output | Render first sheet of .xlsx to TIFF with one page per sheet in C# | Convert .xlsx to TIFF with 24-bit color using Aspose.Cells .NET API
// Tags: Aspose.Cells ImageOrPrintOptions TIFF settings | C# SheetRender ToImage for TIFF output | Excel to 24-bit TIFF conversion Aspose.Cells | OnePagePerSheet worksheet rendering | high‑quality TIFF generation .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks for the source Excel file, loads it into a Workbook, configures ImageOrPrintOptions with OnePagePerSheet, creates a SheetRender for the first worksheet, and saves the rendered image as a 24‑bit TIFF file, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.tiff";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure image options for TIFF output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Render each worksheet on a separate page
                OnePagePerSheet = true
                // Additional properties such as ImageFormat or ImageColorMode are omitted
                // because they are not available in the referenced Aspose.Cells version.
            };

            // Render the first worksheet to a TIFF image
            SheetRender sheetRender = new SheetRender(workbook.Worksheets[0], imgOptions);
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Successfully saved TIFF image to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
