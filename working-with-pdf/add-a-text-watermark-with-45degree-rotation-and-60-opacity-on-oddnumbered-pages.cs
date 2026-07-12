using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // Author: Aspose.Cells .NET example author
    class Program
    {
        static void Main()
        {
            // Create a new workbook with three worksheets (three pages)
            Workbook workbook = new Workbook();
            Worksheet ws1 = workbook.Worksheets[0];
            ws1.Name = "Page1";
            ws1.Cells["A1"].PutValue("Content for page 1");

            int idx = workbook.Worksheets.Add();
            Worksheet ws2 = workbook.Worksheets[idx];
            ws2.Name = "Page2";
            ws2.Cells["A1"].PutValue("Content for page 2");

            idx = workbook.Worksheets.Add();
            Worksheet ws3 = workbook.Worksheets[idx];
            ws3.Name = "Page3";
            ws3.Cells["A1"].PutValue("Content for page 3");

            // Create a rendering font for the watermark text
            RenderingFont font = new RenderingFont("Arial", 36);
            font.Bold = true;
            font.Color = Color.Red;

            // Initialize a text watermark with the desired text and font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font);

            // Configure watermark appearance
            watermark.Rotation = 45f;          // 45‑degree rotation
            watermark.Opacity = 0.6f;         // 60% opacity
            watermark.IsBackground = true;   // Place behind page contents
            watermark.HAlignment = TextAlignmentType.Center;
            watermark.VAlignment = TextAlignmentType.Center;
            watermark.ScaleToPagePercent = 100;

            // Note: Aspose.Cells applies the watermark to all pages via PdfSaveOptions.
            // The current API does not provide a direct way to limit the watermark to odd‑numbered pages.

            // Set the watermark in PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.Watermark = watermark;

            // Save the workbook as PDF with the configured watermark
            workbook.Save("OddPagesWatermark.pdf", pdfOptions);
        }
    }
}