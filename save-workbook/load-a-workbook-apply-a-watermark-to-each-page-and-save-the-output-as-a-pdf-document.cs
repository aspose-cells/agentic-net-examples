using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the existing workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Create a font for the watermark text
        RenderingFont font = new RenderingFont("Calibri", 68)
        {
            Italic = true,
            Bold = true,
            Color = Color.Blue
        };

        // Create a text watermark using the font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,   // Horizontal center
            VAlignment = TextAlignmentType.Center,   // Vertical center
            Rotation = 45,                           // Rotate 45 degrees
            Opacity = 0.3f,                          // 30% opacity
            ScaleToPagePercent = 75,                 // Scale relative to page
            IsBackground = true                      // Place behind page content
        };

        // Configure PDF save options with the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied to each page
        string outputPath = "output_watermark.pdf";
        workbook.Save(outputPath, pdfOptions);
    }
}