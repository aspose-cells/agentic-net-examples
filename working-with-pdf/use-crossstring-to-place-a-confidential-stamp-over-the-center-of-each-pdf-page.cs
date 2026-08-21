// Title: C# – Add a Centered Semi‑Transparent “CONFIDENTIAL” Watermark to PDFs with Aspose.Cells
// Description: Demonstrates how to create a Workbook, define a bold red Arial font, build a RenderingWatermark with the text “CONFIDENTIAL”, center it horizontally and vertically, set 30% opacity, place it over the page content, assign it to PdfSaveOptions, and save the workbook as a PDF where every page shows the stamp.
// Keywords: Aspose.Cells PDF watermark C# | RenderingWatermark example | centered confidential stamp | semi transparent PDF watermark | PdfSaveOptions watermark | Aspose.Cells C# tutorial
// Common Searches: Aspose.Cells add confidential watermark to PDF C# | center text watermark Aspose.Cells PDF | set opacity for PDF watermark using Aspose.Cells | how to place watermark over page content Aspose.Cells | C# code for PDF watermark with Aspose.Cells
// Developer Intent: Apply a centered, semi‑transparent “CONFIDENTIAL” text stamp to each page of a PDF generated from an Aspose.Cells workbook.
// Use Cases: Protect internal reports by overlaying a confidentiality notice. | Mark exported invoices or contracts with a legal disclaimer. | Ensure all pages of multi‑sheet Excel exports carry a uniform security label.
// AI Prompts: Generate C# code that rotates the CONFIDENTIAL watermark 45° diagonally while keeping it centered. | Show how to change the watermark font to Times New Roman, color to blue, and retain the same alignment and opacity. | Explain how to apply the watermark only to selected worksheets when saving the workbook as PDF.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a Workbook, define a bold red Arial font, build a RenderingWatermark with the text “CONFIDENTIAL”, center it horizontally and vertically, set 30% opacity, place it over the page content, assign it to PdfSaveOptions, and save the workbook as a PDF where every page shows the stamp.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add sample content to each worksheet so the PDF has visible pages
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.Cells["A1"].PutValue("Sample content");
        }

        // Define the font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 72)
        {
            Bold = true,
            Color = Color.Red
        };

        // Create a text watermark with the desired stamp
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            // Center the watermark on each page
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            // No rotation (optional: set to 45 for diagonal)
            Rotation = 0,
            // Make the stamp semi‑transparent
            Opacity = 0.3f,
            // Place the watermark over the page content
            IsBackground = false
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the confidential stamp on every page
        workbook.Save("ConfidentialStamped.pdf", pdfOptions);
    }
}
