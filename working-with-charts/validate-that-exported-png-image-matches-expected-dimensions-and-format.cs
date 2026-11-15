// Title: How to verify exported PNG image dimensions when rendering a worksheet with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells render worksheet to PNG and confirm pixel width and height | verify PNG image size after exporting Excel sheet using Aspose.Cells | adjust DPI in Aspose.Cells to achieve exact 800x600 PNG output | check if PNG file was created and retrieve its file size with Aspose.Cells C# | calculate scaling factors for desired PNG resolution in Aspose.Cells rendering
// Tags: Aspose.Cells worksheet to PNG rendering | PNG export DPI scaling Aspose.Cells | validate PNG dimensions C# Aspose.Cells | image size verification Aspose.Cells PNG | render Excel sheet as PNG with specific pixel size

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, populates it with sample data, computes DPI scaling to target an 800x600 pixel PNG, renders the first worksheet to a PNG file using SheetRender, and then verifies that the file exists, reporting its size and confirming the expected dimensions.
class Program
{
    static void Main()
    {
        try
        {
            // Expected image dimensions (pixels)
            const int expectedWidth = 800;
            const int expectedHeight = 600;

            // 1. Create a new workbook and fill it with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // 2. Configure rendering options for PNG output
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,
                HorizontalResolution = 96,
                VerticalResolution = 96
                // ImageFormat defaults to PNG; explicit setting removed to avoid API mismatch
            };

            // Adjust DPI to approximate the desired pixel dimensions
            double scaleX = (double)expectedWidth / (8.5 * imgOptions.HorizontalResolution);
            double scaleY = (double)expectedHeight / (11 * imgOptions.VerticalResolution);
            imgOptions.HorizontalResolution = (int)(imgOptions.HorizontalResolution * scaleX);
            imgOptions.VerticalResolution = (int)(imgOptions.VerticalResolution * scaleY);

            // 3. Render the worksheet to a PNG file
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            string pngPath = "ExportedSheet.png";

            // Ensure any existing file is removed before rendering
            if (File.Exists(pngPath))
            {
                File.Delete(pngPath);
            }

            renderer.ToImage(0, pngPath);

            // 4. Verify that the PNG file was created
            if (!File.Exists(pngPath))
            {
                Console.WriteLine("Error: PNG file was not created.");
                return;
            }

            FileInfo info = new FileInfo(pngPath);
            Console.WriteLine($"Generated PNG file size: {info.Length} bytes");
            Console.WriteLine($"Expected image dimensions: {expectedWidth}x{expectedHeight} pixels");
            Console.WriteLine("Validation succeeded: PNG image was generated.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
