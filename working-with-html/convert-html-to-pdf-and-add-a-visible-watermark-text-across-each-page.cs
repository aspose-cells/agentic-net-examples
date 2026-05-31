using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class HtmlToPdfWithWatermark
{
    static void Main()
    {
        // Load the HTML file into a workbook
        string htmlPath = "input.html";
        Workbook workbook = new Workbook(htmlPath); // HTML is parsed as a workbook

        // Create a font for the watermark text
        RenderingFont watermarkFont = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Color = Color.Red
        };

        // Create a text watermark with desired appearance
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
        {
            HAlignment = TextAlignmentType.Center,   // Horizontal center
            VAlignment = TextAlignmentType.Center,   // Vertical center
            Rotation = 45,                           // Diagonal angle
            Opacity = 0.3f,                          // Semi‑transparent
            ScaleToPagePercent = 70,                 // Size relative to page
            IsBackground = true                      // Render behind page content
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied
        workbook.Save("output.pdf", pdfOptions);
    }
}