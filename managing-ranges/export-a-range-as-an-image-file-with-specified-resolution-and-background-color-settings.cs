// Title: Export Excel range as PNG with specified DPI and background using Aspose.Cells for .NET
// Description: Loads a workbook, selects cells A1:J25 on the first worksheet, configures ImageOrPrintOptions for PNG with 300 dpi horizontal and vertical resolution (white background), converts the range to an image via Range.ToImage, and writes the PNG file to disk.
// Keywords: Aspose.Cells | C# export range to image | Excel range to PNG | custom DPI | background color | Range.ToImage | ImageOrPrintOptions | .NET
// Common Searches: Aspose.Cells export specific range to PNG with 300 DPI | How to set image resolution when converting Excel cells to image in C# | Save Excel range as image with white background using Aspose.Cells | Range.ToImage example with custom DPI | Export Excel table as high‑resolution PNG
// Developer Intent: Generate an image file of a selected cell block with defined resolution and background color settings.
// Use Cases: Insert high‑resolution snapshots of report sections into PDFs or slide decks. | Create thumbnail previews of data tables for web pages or email newsletters. | Provide a static PNG of a worksheet area for documentation without sharing the original workbook.
// AI Prompts: Write C# code that exports a worksheet range to a JPEG image at 150 dpi with a transparent background using Aspose.Cells. | Show how to export multiple non‑contiguous ranges to separate PNG files, each with its own DPI and background color. | Explain how to change the background color of an exported range image when using Aspose.Cells in .NET.

using System;
using System.IO;
using System.Drawing;                     // For Color (kept for potential future use)
using Aspose.Cells;
using Aspose.Cells.Drawing;               // For ImageType
using Aspose.Cells.Rendering;             // For ImageOrPrintOptions

// Loads a workbook, selects cells A1:J25 on the first worksheet, configures ImageOrPrintOptions for PNG with 300 dpi horizontal and vertical resolution (white background), converts the range to an image via Range.ToImage, and writes the PNG file to disk.
class ExportRangeAsImage
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "range_image.png";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the range to export (e.g., A1:J25)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:J25");

            // Configure image options: PNG format, 300 DPI, white background (default)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300
                // BackgroundColor property is not available; PNG default background is white
            };

            // Convert the range to an image (byte array)
            byte[] imageData = range.ToImage(options);

            // Save the image to a file
            File.WriteAllBytes(outputPath, imageData);
            Console.WriteLine($"Range image saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
