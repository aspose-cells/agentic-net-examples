using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class RenderWorkbookToTiff
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample content for 300 DPI TIFF rendering");

        // Configure image options: TIFF format with 300 DPI resolution
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,          // Set output format to TIFF
            HorizontalResolution = 300,          // 300 DPI horizontally
            VerticalResolution = 300,            // 300 DPI vertically
            TiffCompression = TiffCompression.CompressionLZW // Optional compression
        };

        // Render the entire workbook to a TIFF file
        WorkbookRender renderer = new WorkbookRender(workbook, options);
        renderer.ToImage("output_300dpi.tiff");

        Console.WriteLine("Workbook successfully rendered to TIFF at 300 DPI.");
    }
}