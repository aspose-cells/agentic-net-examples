// Title: Add a 30‑point semi‑transparent text watermark to every PDF page with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, inserts sample data, defines a 30‑point bold Arial gray font, builds a centered "CONFIDENTIAL" RenderingWatermark with 40 % opacity placed behind the content, assigns it to PdfSaveOptions.Watermark, and saves the workbook as a PDF where the watermark appears on all pages.
// Keywords: Aspose.Cells PDF watermark C# | RenderingWatermark | 30 point font watermark | 40% opacity watermark | PdfSaveOptions Watermark | add text watermark to PDF | C# Aspose.Cells PDF export | centered background watermark | semi transparent PDF watermark | Aspose.Cells .NET example
// Common Searches: How to add a semi‑transparent text watermark to a PDF using Aspose.Cells C# | Aspose.Cells add 30‑point watermark when saving workbook as PDF | Set opacity and alignment for PDF watermark in Aspose.Cells .NET | Create background watermark for all PDF pages with Aspose.Cells | C# code to apply text watermark to PDF export from Excel
// Developer Intent: Apply a centered 30‑point text watermark with 40 % opacity to every page of a PDF generated from an Aspose.Cells workbook.
// Use Cases: Mark confidential reports with a faint "CONFIDENTIAL" overlay on each page | Brand exported PDFs with a company name as a background watermark | Indicate draft status on legal documents by adding a semi‑transparent "DRAFT" watermark | Provide visual protection for shared spreadsheets converted to PDF
// AI Prompts: Generate code to change the watermark text and color based on worksheet values using Aspose.Cells. | Show how to combine a text watermark with an image watermark, including custom opacity and rotation, in PDF output. | Explain how to apply different watermarks to selected pages when saving a workbook as PDF with Aspose.Cells. | Provide a version of the example that loads an existing workbook instead of creating a new one.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This C# example creates a workbook, inserts sample data, defines a 30‑point bold Arial gray font, builds a centered "CONFIDENTIAL" RenderingWatermark with 40 % opacity placed behind the content, assigns it to PdfSaveOptions.Watermark, and saves the workbook as a PDF where the watermark appears on all pages.
class AddWatermark
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample content for PDF with watermark");

        // Create a rendering font with 30‑point size
        RenderingFont font = new RenderingFont("Arial", 30)
        {
            Bold = true,
            Color = Color.Gray
        };

        // Create a text watermark using the font
        RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
        {
            Opacity = 0.4f,                     // 40% opacity
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,
            IsBackground = true                 // place behind page contents
        };

        // Configure PDF save options with the watermark
        PdfSaveOptions options = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the watermark applied to all pages
        workbook.Save("OutputWithWatermark.pdf", options);
    }
}
