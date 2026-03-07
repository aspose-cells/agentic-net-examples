using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class AddWatermarkToPdf
{
    static void Main()
    {
        // Load the source Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Define the font for the text watermark
        RenderingFont font = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Color = Color.LightGray
        };

        // Create a text watermark with desired properties
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45f,
            Opacity = 0.3f,
            IsBackground = true
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