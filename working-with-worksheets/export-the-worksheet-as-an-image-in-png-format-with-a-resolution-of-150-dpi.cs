// Title: Export a worksheet to a high‑resolution PNG (150 DPI) using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet to PNG with 150 DPI | How to set image resolution when converting Excel sheet to PNG in Aspose.Cells | Aspose.Cells SheetRender ToImage high resolution PNG example | Saving Excel worksheet as PNG at specific DPI using Aspose.Cells .NET | Render entire Excel worksheet to a single PNG page with custom DPI in C#
// Tags: Aspose.Cells worksheet to PNG export | ImageOrPrintOptions DPI setting | SheetRender ToImage high resolution | C# export Excel sheet as PNG | OnePagePerSheet image rendering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, selects a worksheet, configures ImageOrPrintOptions with 150 horizontal and vertical DPI and OnePagePerSheet enabled, then uses SheetRender to render the first page of the sheet as a PNG file, handling missing files and runtime exceptions.
class ExportWorksheetAsPng
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.png";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify by name/index)
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure image export options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Default format is PNG; explicit setting omitted to avoid missing API issues
                HorizontalResolution = 150, // Set horizontal DPI
                VerticalResolution = 150,   // Set vertical DPI
                OnePagePerSheet = true      // Export whole sheet on a single page
            };

            // Render the worksheet to an image
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
            sheetRender.ToImage(0, outputPath); // Export the first page as PNG

            Console.WriteLine($"Worksheet exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
