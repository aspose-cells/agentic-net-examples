// Title: C# – Add a Semi‑Transparent Centered Text Watermark to Every PDF Page with Aspose.Cells
// Description: Demonstrates how to create a workbook, define a large bold Arial font, and apply a RenderingWatermark (text "CONFIDENTIAL", centered, 45° rotation, 30% opacity, background layer, 75% page scale) via PdfSaveOptions. The workbook is saved as a PDF where the watermark appears on each page.
// Keywords: Aspose.Cells | C# | PDF watermark | RenderingWatermark | semi transparent watermark | centered text watermark | rotate watermark | opacity setting | background watermark | PdfSaveOptions | Aspose.Cells example
// Common Searches: Aspose.Cells add text watermark to PDF | C# semi transparent PDF watermark Aspose.Cells | RenderingWatermark rotation opacity example | centered background watermark each PDF page | scale watermark to page size Aspose.Cells
// Developer Intent: Apply a semi‑transparent, centered, rotated text watermark to every page when converting an Aspose.Cells workbook to PDF.
// Use Cases: Confidential reports with a faint "CONFIDENTIAL" overlay on all pages. | Company‑branded PDFs where the logo or name appears as a background watermark. | Draft or review documents marked with a semi‑transparent "DRAFT" label across each page.
// AI Prompts: Show C# code to add a semi‑transparent, rotated text watermark to each PDF page using Aspose.Cells. | Explain the RenderingWatermark properties (Opacity, Rotation, HAlignment, VAlignment, ScaleToPagePercent) for PDF output. | Provide step‑by‑step instructions to set a background watermark when saving a workbook as PDF with Aspose.Cells, including font and positioning options.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // Demonstrates how to create a workbook, define a large bold Arial font, and apply a RenderingWatermark (text "CONFIDENTIAL", centered, 45° rotation, 30% opacity, background layer, 75% page scale) via PdfSaveOptions. The workbook is saved as a PDF where the watermark appears on each page.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample content for PDF with watermark.");

            // Define the font for the watermark text
            RenderingFont font = new RenderingFont("Arial", 72)
            {
                Bold = true,
                Color = Color.Gray   // Light gray for a subtle appearance
            };

            // Create a semi‑transparent text watermark
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                HAlignment = TextAlignmentType.Center,   // Center horizontally
                VAlignment = TextAlignmentType.Center,   // Center vertically
                Rotation = 45,                           // Rotate 45 degrees
                Opacity = 0.3f,                          // 30% opacity (semi‑transparent)
                IsBackground = true,                     // Place behind page contents
                ScaleToPagePercent = 75                  // Scale relative to page size
            };

            // Configure PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied to each page
            workbook.Save("output_watermark.pdf", pdfOptions);
        }
    }
}
