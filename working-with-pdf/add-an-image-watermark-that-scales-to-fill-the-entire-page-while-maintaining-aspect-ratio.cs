// Title: Add a Full‑Page Image Watermark to a PDF with Aspose.Cells (C#)
// Description: Demonstrates how to create a Workbook, load an image, configure a RenderingWatermark to fill the page while preserving aspect ratio, set center alignment, background placement, and 30 % opacity, then save the workbook as a PDF with the image covering the entire page.
// Keywords: Aspose.Cells PDF watermark C# | full page image watermark .NET | RenderingWatermark ScaleToPagePercent | PdfSaveOptions watermark Aspose | centered background watermark C# | maintain aspect ratio watermark PDF | Aspose.Cells image watermark example
// Common Searches: Aspose.Cells add image watermark to PDF | C# full‑page PDF watermark with Aspose | scale watermark to page Aspose.Cells | centered background image in PDF using Aspose.Cells | set opacity for PDF watermark .NET
// Developer Intent: Generate a PDF from a workbook that includes a centered, semi‑transparent image covering the whole page.
// Use Cases: Embedding a corporate logo as a full‑page background on exported reports. | Creating confidential PDFs with a stamped image that fills each page. | Producing marketing brochures where a background image automatically spans the page during conversion.
// AI Prompts: Show how to replace the PNG with a JPEG while keeping the full‑page watermark effect. | Provide code to apply different watermarks with distinct opacities to individual PDF pages. | Explain how to compute ScaleToPagePercent dynamically based on the source image size and page dimensions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to create a Workbook, load an image, configure a RenderingWatermark to fill the page while preserving aspect ratio, set center alignment, background placement, and 30 % opacity, then save the workbook as a PDF with the image covering the entire page.
class ImageWatermarkExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data to demonstrate the watermark effect
        workbook.Worksheets[0].Cells["A1"].PutValue("Sample content with full‑page image watermark");

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
            // Scale the watermark to fill the entire page (maintains aspect ratio)
            ScaleToPagePercent = 100,

            // Center the watermark horizontally and vertically
            HAlignment = TextAlignmentType.Center,
            VAlignment = TextAlignmentType.Center,

            // Place the watermark behind the page contents
            IsBackground = true,

            // Optional: make the watermark semi‑transparent
            Opacity = 0.3f
        };

        // Configure PDF save options to use the watermark
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Watermark = watermark
        };

        // Save the workbook as a PDF with the full‑page image watermark
        workbook.Save("WatermarkedFullPage.pdf", pdfOptions);
    }
}
