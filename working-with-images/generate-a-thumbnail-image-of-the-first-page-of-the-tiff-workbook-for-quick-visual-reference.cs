// Title: Create a PNG thumbnail of the first worksheet page from an Excel file using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells generate thumbnail image of first worksheet page | How to render first sheet of an Excel workbook to PNG with Aspose.Cells | Aspose.Cells ImageOrPrintOptions for low‑resolution workbook preview in .NET | Save Excel worksheet as PNG thumbnail using SheetRender in C# | Create quick visual preview of an Excel file with Aspose.Cells rendering API
// Tags: Aspose.Cells SheetRender PNG thumbnail | ImageOrPrintOptions low‑resolution preview | render first worksheet to image C# | Excel workbook thumbnail generation .NET | save worksheet as PNG using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace ThumbnailGenerator
{
    // The sample loads an Excel workbook, selects the first worksheet, configures ImageOrPrintOptions for a 96 DPI PNG image, and uses SheetRender to render only the first page to a thumbnail file (thumbnail.png), including basic file‑existence checks and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed or pass via command‑line arguments
            string inputPath = "input.xlsx";          // source workbook
            string thumbnailPath = "thumbnail.png";   // final thumbnail image

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.Error.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                using (Workbook workbook = new Workbook(inputPath))
                {
                    // Get the first worksheet (index 0)
                    Worksheet sheet = workbook.Worksheets[0];

                    // Configure rendering options for a low‑resolution PNG thumbnail
                    ImageOrPrintOptions options = new ImageOrPrintOptions
                    {
                        ImageType = ImageType.Png,          // PNG output
                        HorizontalResolution = 96,          // Typical screen DPI
                        VerticalResolution = 96,
                        PageIndex = 0,                      // First page
                        PageCount = 1                       // Render only one page
                    };

                    // Render the first page directly to a PNG file
                    SheetRender sheetRender = new SheetRender(sheet, options);
                    sheetRender.ToImage(0, thumbnailPath);
                    Console.WriteLine($"Thumbnail saved to: {thumbnailPath}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while generating the thumbnail:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
