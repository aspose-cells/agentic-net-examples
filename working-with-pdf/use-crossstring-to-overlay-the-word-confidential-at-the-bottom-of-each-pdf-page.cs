using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

public class ConfidentialWatermarkPdf
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Optional: add some sample data so the PDF has visible content
        sheet.Cells["A1"].PutValue("Sample data for PDF");

        // Create a rendering font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 36);

        // Initialize the watermark with the desired text and font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font);

        // Position the watermark at the bottom‑center of each page
        watermark.HAlignment = TextAlignmentType.Center;   // horizontal center
        watermark.VAlignment = TextAlignmentType.Bottom;   // vertical bottom

        // Adjust appearance: no rotation, light opacity, full page scaling
        watermark.Rotation = 0f;
        watermark.Opacity = 0.2f;               // 20 % opacity
        watermark.ScaleToPagePercent = 100;    // scale to page size

        // Attach the watermark to PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Watermark = watermark;

        // Save the workbook as a PDF with the watermark applied
        workbook.Save("Confidential.pdf", pdfOptions);
    }
}
// Author: Aspose.Cells .NET example – adds a cross‑string “CONFIDENTIAL” watermark at the bottom of each PDF page.