// Title: Convert HTML to PDF with a diagonal semi‑transparent watermark using Aspose.Cells for .NET (C#)
// Description: Loads an HTML file into an Aspose.Cells Workbook, creates a RenderingWatermark with custom font, rotation, opacity and scaling, attaches it to PdfSaveOptions, and saves the workbook as a PDF where the watermark appears behind the content on every page.
// Keywords: Aspose.Cells HTML to PDF | C# PDF watermark | RenderingWatermark example | PdfSaveOptions watermark | diagonal text watermark | semi transparent PDF watermark | Aspose.Cells .NET tutorial | convert HTML to PDF C# | global PDF generation
// Common Searches: how to add a diagonal watermark to PDF with Aspose.Cells | Aspose.Cells convert HTML to PDF with watermark C# | set watermark opacity and rotation in PdfSaveOptions | C# example for RenderingWatermark in Aspose.Cells | add background text watermark to every PDF page
// Developer Intent: Create a PDF from an HTML source and overlay a visible, semi‑transparent diagonal watermark on each page using Aspose.Cells for .NET.
// Use Cases: Generate confidential reports from HTML templates with a "CONFIDENTIAL" watermark across all pages. | Produce branded brochures by stamping the company name as a diagonal watermark on HTML‑based PDFs. | Automate legal document output where every PDF derived from HTML must display a compliance watermark.
// AI Prompts: Write C# code with Aspose.Cells to convert an HTML file to PDF and apply a rotated, semi‑transparent text watermark. | Explain how to adjust watermark size, opacity, and alignment when saving HTML as PDF with Aspose.Cells. | Show how to assign different watermark texts to specific pages while merging multiple HTML sheets into a single PDF.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an HTML file into an Aspose.Cells Workbook, creates a RenderingWatermark with custom font, rotation, opacity and scaling, attaches it to PdfSaveOptions, and saves the workbook as a PDF where the watermark appears behind the content on every page.
class HtmlToPdfWithWatermark
{
    static void Main()
    {
        // Load the HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Define the font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Color = Color.Gray
        };

        // Create a text watermark with desired appearance
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45,               // Rotate to appear across the page
            Opacity = 0.3f,              // Semi‑transparent
            ScaleToPagePercent = 75,     // Scale relative to page size
            IsBackground = true          // Place behind page content
        };

        // Configure PDF save options to include the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied
        workbook.Save("output.pdf", pdfOptions);
    }
}
