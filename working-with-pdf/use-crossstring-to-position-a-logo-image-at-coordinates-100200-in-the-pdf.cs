// Title: Position a logo image at (100,200) in a PDF generated with Aspose.Cells (C#)
// Description: This example creates a workbook, loads a PNG logo, configures a RenderingWatermark with OffsetX = 100 and OffsetY = 200, sets the watermark as a centered background element, applies it via PdfSaveOptions, and saves the workbook as a PDF where the logo appears at the specified coordinates.
// Keywords: Aspose.Cells C# PDF watermark | RenderingWatermark offset X Y | place image at coordinates PDF | logo watermark Aspose.Cells | save workbook as PDF with image | C# Aspose.Cells PDF export logo | set watermark position Aspose.Cells
// Common Searches: Aspose.Cells place logo at specific coordinates in PDF | C# RenderingWatermark offset example | How to add a centered image watermark with Aspose.Cells | Set PDF watermark position using Aspose.Cells | Export workbook to PDF with logo at (100,200)
// Developer Intent: Add a logo image at the (100,200) point coordinates in the PDF produced from a workbook.
// Use Cases: Brand every page of a financial report with a company logo positioned precisely. | Create a semi‑transparent logo watermark for custom invoice PDFs. | Generate PDFs where the logo is offset to align with a predefined template layout.
// AI Prompts: Show how to change the watermark image format and adjust its opacity in the same code. | Provide a snippet that adds multiple RenderingWatermark objects at different coordinates. | Explain how to compute OffsetX/OffsetY values to align a logo to the top‑right corner of a PDF page.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsLogoPdf
{
    // This example creates a workbook, loads a PNG logo, configures a RenderingWatermark with OffsetX = 100 and OffsetY = 200, sets the watermark as a centered background element, applies it via PdfSaveOptions, and saves the workbook as a PDF where the logo appears at the specified coordinates.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (single worksheet)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Load logo image bytes (replace with your actual image path)
            string logoPath = "logo.png";
            if (!File.Exists(logoPath))
            {
                Console.WriteLine($"Image file not found: {logoPath}");
                return;
            }
            byte[] logoBytes = File.ReadAllBytes(logoPath);

            // Create an image watermark using the logo bytes
            RenderingWatermark logoWatermark = new RenderingWatermark(logoBytes)
            {
                // Position the watermark at (100, 200) points (offsets from default alignment)
                OffsetX = 100f,
                OffsetY = 200f,
                // Ensure the watermark is treated as a background element
                IsBackground = true,
                // Center alignment so offsets are applied from the center of the page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Optional: set opacity if you want a semi‑transparent logo
                Opacity = 1.0f
            };

            // Configure PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = logoWatermark
            };

            // Save the workbook as PDF; the logo will appear at the specified coordinates
            string outputPdf = "WorkbookWithLogo.pdf";
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"PDF saved successfully to '{outputPdf}'.");
        }
    }
}
