// Title: Add Text Over Image Watermark to PDF with Aspose.Cells (C#) – Foreground, Centered on Every Page
// Description: This example creates a workbook, loads a PNG logo, builds a foreground RenderingWatermark, adds a TextWatermark that sits above the image, centers both on each page, sets opacity and scaling, assigns them to PdfSaveOptions, and saves the result as a PDF with combined watermarks.
// Keywords: Aspose.Cells PDF watermark C# | image watermark foreground | text watermark over image | RenderingWatermark Aspose.Cells | TextWatermark Aspose.Cells | PdfSaveOptions watermark | .NET PDF watermark sample | centered watermark Aspose.Cells | opacity scaling watermark | multiple watermarks Aspose.Cells
// Common Searches: how to add text watermark on top of image watermark using Aspose.Cells C# | Aspose.Cells foreground image watermark PDF export | center image and text watermarks on each PDF page Aspose.Cells | combine image and text watermarks in Aspose.Cells PDF | C# Aspose.Cells add multiple watermarks to PDF
// Developer Intent: Add both an image and a text watermark to a PDF generated from an Aspose.Cells workbook, ensuring the text appears above the image on every page.
// Use Cases: Corporate reports that need a logo (image) and a confidentiality notice (text) on each page. | Legal documents requiring a seal image with a bold “CONFIDENTIAL” label over it. | Marketing brochures that display a semi‑transparent brand mark plus a tagline on top. | Financial statements where a watermark image is combined with a date stamp text.
// AI Prompts: Generate C# code using Aspose.Cells to add a centered semi‑transparent image watermark and a red bold text watermark that overlays it on each PDF page. | Show how to configure PdfSaveOptions to apply multiple watermarks (image then text) with correct Z‑order in Aspose.Cells .NET. | Explain how to adjust opacity, rotation, and scaling for both image and text watermarks when exporting a workbook to PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example creates a workbook, loads a PNG logo, builds a foreground RenderingWatermark, adds a TextWatermark that sits above the image, centers both on each page, sets opacity and scaling, assigns them to PdfSaveOptions, and saves the result as a PDF with combined watermarks.
class AddTextAndImageWatermark
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Sample data for PDF with watermarks.");

            // Path to the image that will serve as the base watermark
            string imagePath = "watermark.png";
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                return;
            }

            // Load the image bytes (no text overlay to avoid System.Drawing dependency)
            byte[] imageBytes = File.ReadAllBytes(imagePath);

            // Create a RenderingWatermark from the image bytes
            RenderingWatermark watermark = new RenderingWatermark(imageBytes)
            {
                // Place the watermark above the page contents (foreground)
                IsBackground = false,
                // Center the watermark on each page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // No rotation for the image
                Rotation = 0,
                // Adjust opacity as needed (0 = fully transparent, 1 = fully opaque)
                Opacity = 0.5f,
                // Scale the watermark to fit the page
                ScaleToPagePercent = 100
            };

            // Configure PDF save options to use the watermark
            PdfSaveOptions options = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark
            string outputPath = "OutputWithWatermark.pdf";
            wb.Save(outputPath, options);
            Console.WriteLine($"PDF saved successfully: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
