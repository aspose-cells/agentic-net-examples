using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing;

class WatermarkExample
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for PDF");

        // Define the font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 36)
        {
            Color = Color.LightGray,
            Bold = true
        };

        // Create a text watermark
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            // Align to the top‑left corner of the page
            HAlignment = TextAlignmentType.Left,
            VAlignment = TextAlignmentType.Top,
            // Offset to the required coordinates (50, 400)
            OffsetX = 50,
            OffsetY = 400,
            // Optional visual settings
            Rotation = 0,
            Opacity = 0.3f,
            IsBackground = true
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied to every page
        workbook.Save("WatermarkedOutput.pdf", pdfOptions);
    }
}