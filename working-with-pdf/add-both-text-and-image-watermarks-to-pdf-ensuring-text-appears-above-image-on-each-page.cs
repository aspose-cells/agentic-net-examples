// Title: Add Text and Image Watermarks to a PDF from an Aspose.Cells Workbook (C#) – Text Over Image
// Description: This example creates a workbook, inserts sample data, loads a PNG as a byte array, builds a RenderingWatermark for the image, adds a TextWatermark, sets both to overlay the worksheet content, configures alignment, opacity and scaling, assigns them to PdfSaveOptions, and saves the workbook as a PDF where the text watermark appears above the image on every page.
// Keywords: Aspose.Cells PDF text watermark C# | Aspose.Cells image watermark C# | multiple watermarks Aspose.Cells | PdfSaveOptions watermark order | RenderingWatermark example | TextWatermark Aspose.Cells | overlay watermark PDF Aspose.Cells | C# .NET PDF watermark Aspose
// Common Searches: How to add both text and image watermarks to a PDF using Aspose.Cells for .NET | Aspose.Cells C# overlay text watermark on image watermark | PdfSaveOptions add multiple watermarks Aspose.Cells | Render text above image watermark in PDF generated from Excel | Aspose.Cells RenderingWatermark and TextWatermark usage
// Developer Intent: Add a semi‑transparent image watermark and a text watermark to each page of a PDF generated from an Aspose.Cells workbook, ensuring the text appears on top of the image.
// Use Cases: Brand a PDF report with a centered logo (image) and a confidential label (text) on every page. | Create legally protected PDFs where the watermark text (e.g., "Draft") overlays a faint background image. | Generate multi‑page invoices that display a company seal behind the document title.
// AI Prompts: Generate C# code that adds an image watermark and a text watermark to a PDF using Aspose.Cells, with the text rendered above the image. | Explain how to set watermark opacity, scaling, and alignment for both image and text watermarks in PdfSaveOptions. | Show how to combine RenderingWatermark and TextWatermark so the text appears on top of the image when saving a workbook as PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // This example creates a workbook, inserts sample data, loads a PNG as a byte array, builds a RenderingWatermark for the image, adds a TextWatermark, sets both to overlay the worksheet content, configures alignment, opacity and scaling, assigns them to PdfSaveOptions, and saves the workbook as a PDF where the text watermark appears above the image on every page.
    public class WatermarkWithTextAndImage
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample worksheet content");

                // Path to the base image that will be used as the watermark background
                string imagePath = "watermark.png";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Load the image bytes (no drawing operations to avoid System.Drawing dependency)
                byte[] watermarkImageBytes = File.ReadAllBytes(imagePath);

                // Create a RenderingWatermark from the image bytes
                RenderingWatermark watermark = new RenderingWatermark(watermarkImageBytes)
                {
                    // Position the watermark at the center of each page
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center,
                    // Make the watermark appear above the worksheet content
                    IsBackground = false,
                    // Adjust opacity if needed (0 = fully transparent, 1 = fully opaque)
                    Opacity = 0.5f,
                    // Scale the watermark relative to the page size (100 = original size)
                    ScaleToPagePercent = 100
                };

                // Configure PDF save options with the watermark
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Watermark = watermark
                };

                // Save the workbook as a PDF with the watermark
                string outputPath = "Workbook_With_TextAndImage_Watermark.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved with watermark: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            WatermarkWithTextAndImage.Run();
        }
    }
}
