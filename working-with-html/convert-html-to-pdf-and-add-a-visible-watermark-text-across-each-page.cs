// Title: C# – Convert HTML to PDF with a diagonal semi‑transparent watermark using Aspose.Cells
// Description: Demonstrates loading an HTML file into an Aspose.Cells Workbook, configuring a custom RenderingFont, creating a RenderingWatermark (center‑aligned, 45° rotation, 30 % opacity, 75 % page scale) and applying it via PdfSaveOptions so the watermark is rendered behind the content on each PDF page.
// Keywords: Aspose.Cells HTML to PDF | C# PDF watermark example | RenderingWatermark Aspose.Cells | PdfSaveOptions watermark | diagonal text watermark .NET | semi transparent PDF watermark | convert HTML workbook to PDF | Aspose.Cells code sample
// Common Searches: Aspose.Cells add diagonal watermark when saving HTML as PDF | C# convert HTML to PDF with transparent watermark | How to set watermark opacity and rotation in PdfSaveOptions | Render HTML file to PDF with background watermark using Aspose.Cells | Aspose.Cells example for PDF watermarking
// Developer Intent: Produce a PDF from an HTML source and embed a faint, slanted text watermark on every page using Aspose.Cells for .NET.
// Use Cases: Generate confidential reports by converting HTML templates to PDF with a "CONFIDENTIAL" overlay. | Add a draft or review label to HTML‑based invoices before distribution. | Automate batch conversion of web pages to PDF while branding each document with a company watermark.
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook and saves it as a PDF with a custom diagonal watermark, specifying font, opacity, rotation, and background placement. | Explain how to adjust RenderingWatermark size and alignment for different page dimensions in Aspose.Cells. | Provide a loop‑based C# example that processes multiple HTML files, applying the same watermark to each generated PDF.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates loading an HTML file into an Aspose.Cells Workbook, configuring a custom RenderingFont, creating a RenderingWatermark (center‑aligned, 45° rotation, 30 % opacity, 75 % page scale) and applying it via PdfSaveOptions so the watermark is rendered behind the content on each PDF page.
class HtmlToPdfWithWatermark
{
    static void Main()
    {
        // Load the HTML file into a workbook
        string htmlPath = "input.html";
        Workbook workbook = new Workbook(htmlPath);

        // Create a font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 48)
        {
            Bold = true,
            Color = Color.FromArgb(128, 0, 0, 255) // semi‑transparent blue
        };

        // Create a text watermark and configure its appearance
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            Rotation = 45,               // diagonal across the page
            Opacity = 0.3f,              // 30% opacity
            ScaleToPagePercent = 75,     // occupy 75% of page size
            IsBackground = true          // place behind page content
        };

        // Set PDF save options with the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied
        string pdfPath = "output.pdf";
        workbook.Save(pdfPath, pdfOptions);
    }
}
