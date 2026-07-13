using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WatermarkPdfDemo
{
    // Author: Aspose.Cells example – adds a semi‑transparent text watermark to each PDF page
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Optionally add some content to demonstrate multiple pages
            workbook.Worksheets[0].Cells["A1"].PutValue("First Page");
            int sheetIndex = workbook.Worksheets.Add();
            workbook.Worksheets[sheetIndex].Cells["A1"].PutValue("Second Page");
            sheetIndex = workbook.Worksheets.Add();
            workbook.Worksheets[sheetIndex].Cells["A1"].PutValue("Third Page");

            // Create a rendering font for the watermark text
            RenderingFont watermarkFont = new RenderingFont("Arial", 36);
            watermarkFont.Bold = true;
            watermarkFont.Color = Color.Gray; // Light gray for subtle appearance

            // Initialize the watermark with text and the font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont);

            // Configure watermark appearance
            watermark.Rotation = 45f;               // Diagonal across the page
            watermark.Opacity = 0.3f;               // Semi‑transparent
            watermark.IsBackground = true;         // Render behind the sheet content
            watermark.HAlignment = TextAlignmentType.Center;
            watermark.VAlignment = TextAlignmentType.Center;
            watermark.ScaleToPagePercent = 150;     // Slightly larger than page size

            // Set the watermark in PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.Watermark = watermark;

            // Save the workbook as a PDF with the watermark applied
            workbook.Save("Workbook_With_Watermark.pdf", pdfOptions);
        }
    }
}