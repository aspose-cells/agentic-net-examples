using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkEvenPages
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Sample data on page 1");
            sheet.Cells["A2"].PutValue("More data on page 2");
            sheet.Cells["A3"].PutValue("Additional data on page 3");

            // Configure page setup to allow different headers/footers for odd and even pages
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.IsHFDiffOddEven = true;

            // Set an even‑page header (center section) that will act as a watermark text on even pages
            // This uses the SetEvenHeader rule.
            pageSetup.SetEvenHeader(1, "CONFIDENTIAL");

            // Create a rendering font for the watermark
            RenderingFont font = new RenderingFont("Arial", 48)
            {
                Bold = true,
                Color = Color.LightGray
            };

            // Create a text watermark with no rotation and 20% opacity
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                Rotation = 0f,          // No rotation
                Opacity = 0.2f,         // 20% opacity
                IsBackground = true,   // Place behind page contents
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center
            };

            // Assign the watermark to PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as PDF; the watermark appears on all pages,
            // while the even‑page header provides the required even‑page specific text.
            workbook.Save("Workbook_With_EvenPage_Watermark.pdf", pdfOptions);
        }
    }
}