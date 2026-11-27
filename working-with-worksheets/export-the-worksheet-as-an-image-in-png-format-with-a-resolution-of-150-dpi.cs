// Title: How to export an Excel worksheet as a 150 DPI PNG image with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export worksheet to PNG with custom DPI | Set image resolution when rendering an Excel sheet to PNG using Aspose.Cells | C# code to save the first page of a workbook as a high‑resolution PNG file | Change horizontal and vertical resolution for PNG output in Aspose.Cells | Render Excel worksheet to 150 DPI PNG image in .NET
// Tags: custom DPI rendering with ImageOrPrintOptions | SheetRender ToImage PNG export | high‑resolution PNG export from Excel worksheet | Aspose.Cells C# image resolution settings | export first worksheet page as PNG

using System;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing Excel workbook, selects the first worksheet, configures ImageOrPrintOptions to 150 dpi for both horizontal and vertical resolution (PNG is the default format), uses SheetRender to render the first page, and saves it as a PNG file while handling missing files and runtime exceptions.
class ExportWorksheetAsPng
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.png";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure image rendering options (PNG is the default format)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 150,
                VerticalResolution = 150
                // ImageFormat defaults to PNG; explicit setting is optional
            };

            // Render the worksheet to an image
            SheetRender sheetRender = new SheetRender(worksheet, imgOptions);

            // Export the first page (page index 0) as a PNG file
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Worksheet exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
