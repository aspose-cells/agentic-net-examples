// Title: C# – Add a Centered Arial Text Watermark (30 pt, 50 % Opacity) to PDF with Aspose.Cells
// Description: Demonstrates how to create a Workbook, define an Arial 30‑point RenderingFont, build a RenderingWatermark with the text "CONFIDENTIAL", center it horizontally and vertically, set 50 % opacity, place it behind the page content, attach the watermark to PdfSaveOptions, and save the workbook as a PDF that displays the centered watermark.
// Keywords: Aspose.Cells | C# | .NET PDF watermark | centered text watermark | Arial 30pt | opacity 0.5 | RenderingWatermark | PdfSaveOptions | background watermark | workbook to PDF
// Common Searches: Aspose.Cells add centered watermark to PDF | C# set watermark opacity Aspose.Cells | How to use RenderingWatermark with Arial font | Save workbook as PDF with background text watermark | Center text watermark in PDF using Aspose.Cells .NET
// Developer Intent: Generate a PDF from a workbook that includes a centered, semi‑transparent Arial text watermark.
// Use Cases: Mark confidential reports with a discreet background label before sharing. | Brand marketing PDFs by embedding the company name as a centered watermark. | Apply a legal disclaimer watermark to invoices to satisfy compliance rules.
// AI Prompts: Write C# code with Aspose.Cells to add a diagonal red watermark at 70 % opacity to a PDF. | Show how to assign different watermarks to each worksheet when exporting them as separate PDFs. | Explain how to calculate watermark position dynamically based on page dimensions in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // Demonstrates how to create a Workbook, define an Arial 30‑point RenderingFont, build a RenderingWatermark with the text "CONFIDENTIAL", center it horizontally and vertically, set 50 % opacity, place it behind the page content, attach the watermark to PdfSaveOptions, and save the workbook as a PDF that displays the centered watermark.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Optional: add some sample data to visualize the watermark effect
            sheet.Cells["A1"].PutValue("Sample content for watermark demonstration.");

            // Create a RenderingFont with Arial, 30‑point size
            RenderingFont font = new RenderingFont("Arial", 30);

            // Create a text watermark using the font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                // Center the watermark horizontally and vertically
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,

                // Set opacity to 50%
                Opacity = 0.5f,

                // Place the watermark behind the page contents (optional)
                IsBackground = true
            };

            // Configure PDF save options with the watermark
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as PDF with the centered watermark
            workbook.Save("CenteredWatermark.pdf", saveOptions);
        }
    }
}
