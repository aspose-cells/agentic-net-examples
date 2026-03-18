using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to demonstrate the watermark effect
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample data for watermark demo");

        // Create a rendering font that defines the appearance of the watermark text
        RenderingFont font = new RenderingFont("Calibri", 68)
        {
            Italic = true,
            Bold = true,
            Color = Color.Blue
        };

        // Create a text watermark using the font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,   // Horizontal center alignment
            VAlignment = TextAlignmentType.Center,   // Vertical center alignment
            Rotation = 45,                           // Rotate 45 degrees
            Opacity = 0.3f,                          // 30% opacity
            ScaleToPagePercent = 75,                 // Scale relative to page size
            IsBackground = true                      // Place behind page contents
        };

        // Configure PDF save options and assign the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF file with the applied watermark
        workbook.Save("output_watermark.pdf", pdfOptions);
    }
}