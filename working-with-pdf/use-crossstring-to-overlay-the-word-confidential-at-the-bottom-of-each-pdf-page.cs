using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample content
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for PDF");

        // Define the font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 36)
        {
            Bold = true,
            Color = Color.Red
        };

        // Create a text watermark with the word "CONFIDENTIAL"
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            // Center horizontally, align to the bottom of each page
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Bottom,
            // Slight offset from the bottom edge (optional)
            OffsetY = 20,
            // Make the watermark semi‑transparent and overlay on top of content
            Opacity = 0.3f,
            IsBackground = false
        };

        // Configure PDF save options to use the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied to every page
        workbook.Save("ConfidentialBottom.pdf", pdfOptions);
    }
}