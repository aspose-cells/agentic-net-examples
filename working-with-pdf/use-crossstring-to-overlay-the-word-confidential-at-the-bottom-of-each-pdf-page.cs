// Title: Add a Bottom‑Centered Semi‑Transparent “CONFIDENTIAL” Watermark to PDF with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, inserts sample data, defines a red 36‑pt Arial font, and applies a RenderingWatermark with the text “CONFIDENTIAL”. The watermark is centered horizontally, aligned to the bottom of each page, set to 30 % opacity, and rendered on top of the content. The configured PdfSaveOptions save the workbook as a PDF where every page displays the overlay watermark.
// Keywords: Aspose.Cells PDF watermark | C# add confidential watermark | bottom centered watermark Aspose | semi transparent PDF watermark | RenderingWatermark example | PdfSaveOptions watermark overlay
// Common Searches: how to add a confidential watermark to each PDF page using Aspose.Cells | Aspose.Cells C# overlay text watermark at bottom of PDF | set watermark opacity and alignment in PdfSaveOptions | render watermark on top of PDF content Aspose.Cells | bottom‑center watermark for PDF export .NET
// Developer Intent: Apply a semi‑transparent “CONFIDENTIAL” text overlay at the bottom of every page when exporting a workbook to PDF.
// Use Cases: Distribute internal reports that must be marked confidential on each page. | Create legal or compliance documents with a visible disclaimer without altering layout. | Generate marketing PDFs that require a bottom‑center notice while preserving original design.
// AI Prompts: Show how to change the watermark text, font size, or color programmatically before saving the PDF with Aspose.Cells. | Provide code to apply different watermarks to odd and even pages during PDF export. | Explain how to toggle between background and overlay watermark modes and adjust opacity in PdfSaveOptions.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example creates a workbook, inserts sample data, defines a red 36‑pt Arial font, and applies a RenderingWatermark with the text “CONFIDENTIAL”. The watermark is centered horizontally, aligned to the bottom of each page, set to 30 % opacity, and rendered on top of the content. The configured PdfSaveOptions save the workbook as a PDF where every page displays the overlay watermark.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample content for PDF");

        // Define the font for the watermark text
        RenderingFont font = new RenderingFont("Arial", 36)
        {
            Bold = true,
            Color = Color.Red
        };

        // Create a text watermark with the word "CONFIDENTIAL"
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            // Center horizontally, align to the bottom of each page
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Bottom,
            // Make the watermark semi‑transparent and render on top of page contents
            Opacity = 0.3f,
            IsBackground = false
        };

        // Configure PDF save options to use the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied to every page
        workbook.Save("ConfidentialOutput.pdf", pdfOptions);
    }
}
