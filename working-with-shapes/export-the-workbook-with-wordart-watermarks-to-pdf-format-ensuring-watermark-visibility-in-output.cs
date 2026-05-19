using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data so the worksheet has visible content
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue("Row 2");
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Create a rendering font for the WordArt style watermark
            RenderingFont font = new RenderingFont("Calibri", 68)
            {
                Italic = true,
                Bold = true,
                Color = Color.Blue
            };

            // Create a text watermark using the font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                // Center the watermark on the page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Rotate for a typical WordArt effect
                Rotation = 45,
                // Make it semi‑transparent so underlying content remains readable
                Opacity = 0.3f,
                // Scale relative to the page size
                ScaleToPagePercent = 75,
                // Place the watermark behind the worksheet content
                IsBackground = true
            };

            // Configure PDF save options and assign the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied
            workbook.Save("WorkbookWithWatermark.pdf", pdfOptions);
        }
    }
}