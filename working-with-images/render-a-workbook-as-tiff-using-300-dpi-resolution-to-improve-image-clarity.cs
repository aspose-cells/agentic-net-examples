using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class RenderWorkbookToTiff
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for TIFF rendering");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Configure image options: TIFF format with 300 DPI resolution
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Tiff;          // Set output format to TIFF
        options.HorizontalResolution = 300;         // 300 DPI horizontally
        options.VerticalResolution = 300;           // 300 DPI vertically
        options.TiffCompression = TiffCompression.CompressionLZW; // Optional compression

        // Create a renderer for the whole workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render the workbook to a TIFF file
        string outputFile = "workbook_300dpi.tiff";
        renderer.ToImage(outputFile);

        Console.WriteLine($"Workbook successfully rendered to TIFF at {outputFile} with 300 DPI.");
    }
}