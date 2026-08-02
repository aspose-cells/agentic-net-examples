// Title: Add Full‑Page Image Watermark to PDF with Aspise.Cells (C#)
// Description: Creates a workbook, loads a PNG, configures a RenderingWatermark with 100% ScaleToPagePercent, centered alignment, background placement, 30% opacity, and saves the workbook as a PDF where the image fills the page while preserving its aspect ratio.
// Keywords: Aspose.Cells PDF watermark | C# image watermark full page | RenderingWatermark ScaleToPagePercent | Aspose.Cells background image PDF | Excel to PDF watermark Aspose
// Common Searches: Aspose.Cells add image watermark to PDF | scale watermark to fill PDF page .NET | center background watermark Aspose.Cells | set watermark opacity in PDF using Aspose.Cells | full‑page image watermark Excel to PDF
// Developer Intent: Apply an image watermark that covers the entire PDF page, maintains the original aspect ratio, and appears behind the worksheet content.
// Use Cases: Brand every exported PDF report with a semi‑transparent company logo that spans the whole page. | Create printable forms or letterheads where a background image automatically scales to the page size. | Generate invoices or statements with a centered watermark that does not interfere with data readability.
// AI Prompts: Show how to rotate the RenderingWatermark 45° while keeping it scaled to the page. | Provide C# code that adds a text watermark in addition to the image watermark using PdfSaveOptions. | Explain how to compute the exact ScaleToPagePercent value for a custom image size and a specific PDF page dimension.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // Creates a workbook, loads a PNG, configures a RenderingWatermark with 100% ScaleToPagePercent, centered alignment, background placement, 30% opacity, and saves the workbook as a PDF where the image fills the page while preserving its aspect ratio.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample worksheet with image watermark");

            // Load the watermark image into a byte array
            string imagePath = "watermark.png";
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }
            byte[] imageData = File.ReadAllBytes(imagePath);

            // Create an image‑based RenderingWatermark
            RenderingWatermark watermark = new RenderingWatermark(imageData)
            {
                // Scale to fill the entire page (maintains aspect ratio)
                ScaleToPagePercent = 100,

                // Center the watermark on the page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,

                // Place the watermark behind the page content
                IsBackground = true,

                // Optional: set opacity and rotation
                Opacity = 0.3f,
                Rotation = 0f
            };

            // Configure PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the image watermark
            workbook.Save("WatermarkedOutput.pdf", pdfOptions);

            Console.WriteLine("PDF saved with image watermark.");
        }
    }
}
