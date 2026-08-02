// Title: Add an Image Watermark (original size) to PDF Export in Aspose.Cells C#
// Description: Demonstrates how to load an image and apply it as a watermark to a PDF generated from an Aspose.Cells workbook. The watermark keeps its original dimensions, is placed in the lower‑right corner, and can be rendered with custom opacity and layering.
// Keywords: Aspose.Cells PDF watermark C# | image watermark lower right | keep original image size watermark | RenderingWatermark Aspose.Cells | PDF export with watermark | C# Aspose.Cells example
// Common Searches: Aspose.Cells add image watermark to PDF | C# place watermark bottom right in PDF | keep original size watermark Aspose.Cells | set opacity for PDF watermark Aspose.Cells | RenderingWatermark alignment options
// Developer Intent: Apply an unchanged‑size image watermark to the lower‑right corner of a PDF created from an Aspose.Cells workbook.
// Use Cases: Brand a report PDF with a logo positioned at the bottom‑right without scaling. | Add a semi‑transparent signature image to each exported PDF page while preserving its resolution. | Overlay a custom graphic in front of worksheet content during PDF conversion.
// AI Prompts: Show how to rotate the RenderingWatermark 45° while keeping it anchored to the lower‑right corner. | Provide C# code to add multiple image watermarks to different corners of a PDF using Aspose.Cells. | Explain how to adjust watermark opacity dynamically based on page content during PDF export.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsWatermarkDemo
{
    // Demonstrates how to load an image and apply it as a watermark to a PDF generated from an Aspose.Cells workbook. The watermark keeps its original dimensions, is placed in the lower‑right corner, and can be rendered with custom opacity and layering.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (empty workbook with a default worksheet)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Optional: put some sample data to visualize the watermark position
            sheet.Cells["A1"].PutValue("Sample data in the worksheet.");

            // Load the image that will be used as the watermark
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
                // Align to the lower‑right corner of the page
                HAlignment = TextAlignmentType.Right,
                VAlignment = TextAlignmentType.Bottom,

                // Keep the original image size (scale to 100% of the page size)
                ScaleToPagePercent = 100,

                // Optional visual settings
                Opacity = 0.5f,          // 50% transparent
                Rotation = 0,           // No rotation
                IsBackground = false    // Place watermark in front of page contents
            };

            // Configure PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as PDF with the watermark applied
            string outputPath = "WorkbookWithImageWatermark.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved with image watermark at: {outputPath}");
        }
    }
}
