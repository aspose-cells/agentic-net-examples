// Title: C# – Convert Excel to PDF with a Semi‑Transparent Diagonal Watermark using Aspose.Cells
// Description: A concise C# example that loads an .xlsx workbook, builds a RenderingFont and a RenderingWatermark (centered, 45° rotation, 30% opacity, placed behind content), attaches the watermark to PdfSaveOptions, and saves the workbook as a PDF where the watermark appears on every page.
// Keywords: Aspose.Cells | C# | .NET | Excel to PDF conversion | PDF watermark | RenderingWatermark | PdfSaveOptions | semi transparent watermark | diagonal watermark | background watermark | code sample | GitHub
// Common Searches: How to add a diagonal semi‑transparent watermark when saving Excel to PDF with Aspose.Cells | Aspose.Cells C# example for PDF conversion with watermark | Set opacity for PDF watermark using RenderingWatermark | Add text watermark behind content in PDF generated from Excel | Aspose.Cells PDFSaveOptions watermark property
// Developer Intent: The developer needs to convert an Excel workbook to PDF and embed a semi‑transparent diagonal watermark on each page using Aspose.Cells for .NET.
// Use Cases: Create confidential reports by converting internal spreadsheets to PDF with a "CONFIDENTIAL" watermark. | Automate batch conversion of workbooks to PDF while applying a corporate branding watermark. | Generate legally compliant documents that require a faint background watermark for authenticity.
// AI Prompts: Show how to modify the watermark text, font, color, rotation, and opacity in the Aspose.Cells PDF conversion example. | Provide a C# snippet that adds an image watermark instead of text when saving an Excel workbook to PDF with Aspose.Cells. | Explain how to apply different watermarks to individual worksheets during a multi‑sheet PDF export using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // A concise C# example that loads an .xlsx workbook, builds a RenderingFont and a RenderingWatermark (centered, 45° rotation, 30% opacity, placed behind content), attaches the watermark to PdfSaveOptions, and saves the workbook as a PDF where the watermark appears on every page.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your source file path)
            string sourcePath = "input.xlsx";
            Workbook workbook = new Workbook(sourcePath); // Load rule

            // Create a font for the watermark text
            RenderingFont font = new RenderingFont("Calibri", 68)
            {
                Bold = true,
                Italic = true,
                Color = Color.Blue
            };

            // Create a semi‑transparent text watermark
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                // Center the watermark on each page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Rotate for a diagonal appearance
                Rotation = 45f,
                // Set opacity (0 = fully transparent, 1 = fully opaque)
                Opacity = 0.3f,
                // Scale relative to the page size
                ScaleToPagePercent = 75,
                // Place the watermark behind the page content
                IsBackground = true
            };

            // Configure PDF save options and assign the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied
            string outputPath = "output_watermark.pdf";
            workbook.Save(outputPath, pdfOptions); // Save rule

            Console.WriteLine($"Workbook saved to PDF with watermark: {outputPath}");
        }
    }
}
