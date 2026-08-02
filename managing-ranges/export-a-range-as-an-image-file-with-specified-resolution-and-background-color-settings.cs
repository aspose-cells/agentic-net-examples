// Title: C# – Export Excel Range A1:J25 to PNG with 300 DPI and Custom Background using Aspose.Cells
// Description: Loads an XLSX workbook, creates the A1:J25 range on the first worksheet, configures ImageOrPrintOptions (PNG, 300 DPI, optional background color), converts the range to an image byte array with ToImage, and writes the result to a PNG file.
// Keywords: Aspose.Cells | C# export range to image | range.ToImage | ImageOrPrintOptions | PNG 300 DPI | custom background color | Excel to PNG .NET | high‑resolution Excel screenshot | save range as image | Aspose.Cells rendering
// Common Searches: Aspose.Cells export Excel range as PNG | Set DPI for Excel range image in C# | How to change background color of exported Excel image | range.ToImage Aspose.Cells example | Export specific cells to high‑resolution image .NET
// Developer Intent: Generate a 300 DPI PNG image of cells A1:J25 from an Excel workbook, optionally specifying a background color, using Aspose.Cells for .NET.
// Use Cases: Create high‑resolution screenshots of report sections for documentation or presentations. | Generate thumbnail previews of spreadsheet areas for web dashboards or galleries. | Produce printable images of data blocks with a consistent background for PDF reports.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet range to a JPEG image at 150 DPI and a light‑gray background. | Show how to stream the image byte array returned by range.ToImage directly to an ASP.NET Core HTTP response.

using System;
using System.IO;
using System.Drawing;                     // For Color
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an XLSX workbook, creates the A1:J25 range on the first worksheet, configures ImageOrPrintOptions (PNG, 300 DPI, optional background color), converts the range to an image byte array with ToImage, and writes the result to a PNG file.
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

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create the range A1:J25 that we want to export
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:J25");

            // Configure image conversion options (resolution and background color)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png, // PNG format
                HorizontalResolution = 300,                     // 300 DPI horizontally
                VerticalResolution = 300                        // 300 DPI vertically
                // BackgroundColor defaults to white; omitted if property unavailable
            };

            // Convert the range to an image byte array
            byte[] imageData = range.ToImage(options);

            // Save the resulting image to a file
            File.WriteAllBytes(outputPath, imageData);

            Console.WriteLine("Range exported to image successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
