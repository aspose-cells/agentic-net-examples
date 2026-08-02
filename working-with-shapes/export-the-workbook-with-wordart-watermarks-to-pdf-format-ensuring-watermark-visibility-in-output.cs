// Title: Export Aspose.Cells Workbook to PDF with a WordArt‑style Text Watermark (C#)
// Description: Creates a new workbook, adds sample data, defines a large italic‑bold blue font, builds a RenderingWatermark with the text "CONFIDENTIAL" centered, rotated 45°, semi‑transparent, and placed on top of the sheet, then saves the workbook as a PDF using PdfSaveOptions.
// Keywords: Aspose.Cells PDF watermark C# | RenderingWatermark example | WordArt text watermark Aspose.Cells | PdfSaveOptions watermark | export workbook to PDF with watermark
// Common Searches: Aspose.Cells add rotated text watermark to PDF | C# export workbook to PDF with visible watermark | How to set opacity and alignment for Aspose.Cells PDF watermark | Create WordArt‑style watermark in Aspose.Cells PDF output
// Developer Intent: Apply a WordArt‑style text watermark to a workbook and generate a PDF where the watermark overlays the worksheet content.
// Use Cases: Confidential reports that need a prominent overlay label. | Branded PDFs with a semi‑transparent company name displayed diagonally. | Legal documents marked with "CONFIDENTIAL" to discourage unauthorized sharing.
// AI Prompts: Show how to move the watermark behind the worksheet data instead of on top. | Provide a sample that uses an image file as a watermark with RenderingWatermark. | Explain how to calculate watermark opacity based on page size or DPI.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkPdf
{
    // Creates a new workbook, adds sample data, defines a large italic‑bold blue font, builds a RenderingWatermark with the text "CONFIDENTIAL" centered, rotated 45°, semi‑transparent, and placed on top of the sheet, then saves the workbook as a PDF using PdfSaveOptions.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data to make the watermark visible over content
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue("Row 2");
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Create a rendering font for the watermark text
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
                // Semi‑transparent so underlying data is still readable
                Opacity = 0.3f,
                // Scale relative to the page size
                ScaleToPagePercent = 75,
                // Place the watermark on top of the content to ensure visibility
                IsBackground = false
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
