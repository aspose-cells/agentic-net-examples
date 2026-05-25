using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class AddCenteredWatermark
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Optional: add some sample content to the worksheet
        sheet.Cells["A1"].PutValue("Sample data");

        // Define the watermark font (Arial, 30‑point)
        RenderingFont font = new RenderingFont("Arial", 30)
        {
            // Light gray makes the watermark visible but not intrusive
            Color = Color.LightGray
        };

        // Create a text watermark with the specified font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            // Center the watermark horizontally and vertically
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            // Set opacity to 50%
            Opacity = 0.5f,
            // Place the watermark behind the page content
            IsBackground = true
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the centered watermark
        workbook.Save("CenteredWatermark.pdf", pdfOptions);
    }
}