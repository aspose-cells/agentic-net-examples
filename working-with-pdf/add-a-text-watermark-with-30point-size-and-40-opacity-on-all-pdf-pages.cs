// Title: Add a 30‑pt Gray Text Watermark (40% Opacity) to All PDF Pages with Aspose.Cells for .NET
// Description: This example creates a workbook, inserts sample data, defines a 30‑point bold Arial gray font, and builds a RenderingWatermark with the text "CONFIDENTIAL". The watermark is centered, placed behind the content, set to 40% opacity, and attached to PdfSaveOptions before saving the workbook as a PDF, applying the mark to every page.
// Keywords: Aspose.Cells PDF watermark C# | RenderingWatermark example | 30‑point text watermark | watermark opacity 40 percent | centered PDF watermark Aspose | PdfSaveOptions watermark | C# add text watermark to PDF
// Common Searches: how to add a text watermark with Aspose.Cells .NET | C# set watermark opacity when saving workbook to PDF | centered gray watermark on all PDF pages Aspose | apply background watermark using PdfSaveOptions | Aspose.Cells render watermark on each page
// Developer Intent: Generate a PDF from a workbook that includes a centered 30‑pt gray text watermark with 40% opacity on every page.
// Use Cases: Issue confidential reports where "CONFIDENTIAL" appears behind the content. | Brand PDFs with a semi‑transparent company name across all pages. | Mark draft documents with a light watermark to indicate they are not final.
// AI Prompts: Show how to rotate the watermark 45° and change its color to red while keeping 40% opacity. | Demonstrate applying different watermarks to individual worksheets before exporting each as a separate PDF.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkExample
{
    // This example creates a workbook, inserts sample data, defines a 30‑point bold Arial gray font, and builds a RenderingWatermark with the text "CONFIDENTIAL". The watermark is centered, placed behind the content, set to 40% opacity, and attached to PdfSaveOptions before saving the workbook as a PDF, applying the mark to every page.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty or load an existing one)
            Workbook workbook = new Workbook();

            // Add some sample content to demonstrate the watermark effect
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data with watermark");

            // Create a RenderingFont with 30‑point size
            RenderingFont font = new RenderingFont("Arial", 30)
            {
                Bold = true,
                Color = Color.Gray
            };

            // Create a text watermark using the font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                // Set watermark properties
                Opacity = 0.4f,                     // 40% opacity
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                IsBackground = true,               // place behind page contents
                Rotation = 0                        // no rotation
            };

            // Configure PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as PDF with the watermark applied to all pages
            workbook.Save("WatermarkedOutput.pdf", pdfOptions);
        }
    }
}
