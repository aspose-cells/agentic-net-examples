// Title: Add a semi‑transparent diagonal CONFIDENTIAL watermark to PDF pages with Aspose.Cells
// Description: Demonstrates creating a workbook, defining a light‑gray Arial font, configuring a RenderingWatermark (centered, 45° rotation, 30% opacity, 75% page scale, background placement) and applying it via PdfSaveOptions.Watermark to produce a PDF where every page shows the semi‑transparent CONFIDENTIAL watermark.
// Keywords: Aspose.Cells PDF watermark | PdfSaveOptions Watermark | RenderingWatermark C# | semi transparent text watermark | diagonal PDF watermark Aspose | C# Aspose.Cells export PDF | background watermark workbook
// Common Searches: Aspose.Cells add watermark to PDF | C# set text watermark opacity Aspose.Cells | PdfSaveOptions watermark rotation example | How to create diagonal watermark in PDF with Aspose.Cells | RenderingWatermark background placement
// Developer Intent: Embed a semi‑transparent diagonal CONFIDENTIAL text watermark on every page of a PDF generated from an Aspose.Cells workbook.
// Use Cases: Confidential internal reports with a light‑gray watermark behind data | Legal or compliance documents that need a visible yet unobtrusive watermark | Brand‑protected marketing PDFs where the company name appears as a background watermark | Audit‑ready spreadsheets exported to PDF with a tamper‑evident label
// AI Prompts: Show how to change the watermark text, font size, and color dynamically based on worksheet values in Aspose.Cells. | Provide code to add multiple watermarks with different rotations on the same PDF page using PdfSaveOptions. | Explain how to adjust watermark opacity and scaling for varying page sizes when exporting to PDF.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates creating a workbook, defining a light‑gray Arial font, configuring a RenderingWatermark (centered, 45° rotation, 30% opacity, 75% page scale, background placement) and applying it via PdfSaveOptions.Watermark to produce a PDF where every page shows the semi‑transparent CONFIDENTIAL watermark.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample content for PDF with watermark");

        // Create a rendering font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 72)
        {
            Bold = true,
            Color = Color.LightGray   // light color for semi‑transparent effect
        };

        // Create a text watermark and configure its appearance
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            HAlignment = TextAlignmentType.Center,   // horizontal center
            VAlignment = TextAlignmentType.Center,   // vertical center
            Rotation = 45f,                          // diagonal across the page
            Opacity = 0.3f,                          // semi‑transparent
            ScaleToPagePercent = 75,                 // size relative to page
            IsBackground = true                      // place behind page contents
        };

        // Set the watermark in PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the configured watermark
        workbook.Save("output_watermark.pdf", pdfOptions);
    }
}
