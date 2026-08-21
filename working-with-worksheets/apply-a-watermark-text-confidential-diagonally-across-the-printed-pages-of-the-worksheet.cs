// Title: Add a Diagonal “Confidential” Text Watermark to PDF Exported from Aspose.Cells (.NET)
// Description: Demonstrates how to create a workbook, define a RenderingFont and RenderingWatermark with the text "Confidential", set a 45° rotation, 30% opacity, background placement, centered alignment, and full‑page scaling, then assign the watermark to PdfSaveOptions and save the worksheet as a PDF where the watermark appears on every printed page.
// Keywords: Aspose.Cells PDF watermark | C# diagonal text watermark | RenderingWatermark rotation | PDFSaveOptions background watermark | Aspose.Cells .NET example | confidential watermark worksheet
// Common Searches: Aspose.Cells add diagonal watermark to PDF | How to set RenderingWatermark opacity in C# | Save worksheet as PDF with background text using Aspose.Cells | C# code for confidential watermark on exported Excel
// Developer Intent: Apply a semi‑transparent diagonal "Confidential" watermark to every page of a PDF generated from an Aspose.Cells worksheet.
// Use Cases: Protect sensitive reports by embedding a discreet diagonal watermark before distribution. | Mark draft or internal documents with a light‑gray background label without altering cell data. | Add branding or legal notices to printed worksheets while keeping the original workbook unchanged.
// AI Prompts: Generate C# code that uses Aspose.Cells to add a semi‑transparent diagonal watermark with custom text and export the worksheet as a PDF. | Explain the RenderingWatermark properties needed to rotate, set opacity, align, and scale a text watermark for PDF output in Aspose.Cells. | Provide step‑by‑step instructions to apply a background watermark only to printed pages while preserving worksheet content.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // Demonstrates how to create a workbook, define a RenderingFont and RenderingWatermark with the text "Confidential", set a 45° rotation, 30% opacity, background placement, centered alignment, and full‑page scaling, then assign the watermark to PdfSaveOptions and save the worksheet as a PDF where the watermark appears on every printed page.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data to visualize the watermark effect
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue("This worksheet will have a diagonal \"Confidential\" watermark.");

            // Create a rendering font for the watermark text
            RenderingFont font = new RenderingFont("Arial", 48)
            {
                Bold = true,
                Color = Color.LightGray
            };

            // Create a text watermark with the desired text and font
            RenderingWatermark watermark = new RenderingWatermark("Confidential", font)
            {
                // Position the watermark diagonally across the page
                Rotation = 45f,
                // Make the watermark semi‑transparent
                Opacity = 0.3f,
                // Place it behind the worksheet content
                IsBackground = true,
                // Center it horizontally and vertically
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Scale it to cover most of the page
                ScaleToPagePercent = 100
            };

            // Configure PDF save options and assign the watermark
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF; the watermark will appear on every printed page
            workbook.Save("ConfidentialWatermarked.pdf", saveOptions);

            Console.WriteLine("Workbook saved with diagonal \"Confidential\" watermark.");
        }
    }
}
