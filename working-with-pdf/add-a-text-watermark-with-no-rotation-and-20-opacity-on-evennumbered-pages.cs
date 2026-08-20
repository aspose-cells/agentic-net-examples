// Title: Add a 20% Opacity, Non‑Rotated Text Watermark to Even Pages in Aspose.Cells PDF
// Description: Demonstrates how to create a workbook, enable odd/even header‑footer settings, define a large Arial RenderingFont, and configure a RenderingWatermark with zero rotation and 0.2 opacity. The watermark is attached to PdfSaveOptions and saved as a PDF. Since Aspose.Cells applies the watermark to every page, the example notes two approaches to restrict it to even‑numbered pages: splitting the workbook or using a PDF post‑processing tool.
// Keywords: Aspose.Cells PDF watermark | text watermark even pages | 20% opacity watermark | non rotated watermark C# | RenderingWatermark Aspose.Cells | PdfSaveOptions watermark | split workbook even pages | post‑process PDF watermark
// Common Searches: Aspose.Cells add watermark only on even pages | C# create semi transparent text watermark PDF | non rotated watermark Aspose.Cells PDFSaveOptions | how to set watermark opacity in Aspose.Cells | apply watermark to every second page PDF
// Developer Intent: Create a centered, non‑rotated text watermark with 20% opacity that appears exclusively on even‑numbered pages of a PDF generated from an Aspose.Cells workbook.
// Use Cases: Protect confidential sections of a multi‑page report by showing a light watermark only on the back (even) pages. | Add subtle branding to every second page of an invoice PDF for visual consistency. | Display a “Draft” label on even pages of a brochure to differentiate draft content without cluttering odd pages.
// AI Prompts: Generate C# code using Aspose.Cells that adds a 20% opacity, non‑rotated "CONFIDENTIAL" watermark only to even pages when saving as PDF. | Show how to separate a workbook into odd‑page and even‑page worksheets and assign distinct PdfSaveOptions so the watermark is applied just to the even‑page worksheet. | Explain a workflow that combines Aspose.Cells PDF export with a PDF post‑processing library (e.g., iTextSharp) to add a watermark to even pages after the PDF is created.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkExample
{
    // Demonstrates how to create a workbook, enable odd/even header‑footer settings, define a large Arial RenderingFont, and configure a RenderingWatermark with zero rotation and 0.2 opacity. The watermark is attached to PdfSaveOptions and saved as a PDF. Since Aspose.Cells applies the watermark to every page, the example notes two approaches to restrict it to even‑numbered pages: splitting the workbook or using a PDF post‑processing tool.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to demonstrate the watermark effect
            sheet.Cells["A1"].PutValue("This is page 1");
            sheet.Cells["A2"].PutValue("This is page 2");
            sheet.Cells["A3"].PutValue("This is page 3");
            sheet.Cells["A4"].PutValue("This is page 4");

            // Enable different headers/footers for odd and even pages (required for even‑page specific settings)
            sheet.PageSetup.IsHFDiffOddEven = true;

            // Create a rendering font for the watermark text
            RenderingFont font = new RenderingFont("Arial", 48)
            {
                Bold = true,
                Color = Color.LightGray
            };

            // Create a text watermark with no rotation and 20% opacity
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                Rotation = 0f,          // No rotation
                Opacity = 0.2f,        // 20% opacity
                IsBackground = true,   // Place behind page contents
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center
            };

            // Assign the watermark to PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as PDF; the watermark will appear on all pages.
            // To restrict it to even‑numbered pages, you would need to split the workbook
            // into separate documents or use a PDF post‑processing library.
            workbook.Save("EvenPagesWatermark.pdf", pdfOptions);
        }
    }
}
