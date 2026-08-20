// Title: C# – Add an Image Watermark to Aspose.Cells PDF (lower‑right, original size)
// Description: Creates an in‑memory workbook, loads a PNG image, configures a RenderingWatermark with right and bottom alignment, 100 % scale, no rotation, full opacity, and applies it via PdfSaveOptions to produce a PDF where the watermark appears in the lower‑right corner at its original dimensions.
// Keywords: Aspose.Cells PDF watermark C# | RenderingWatermark image alignment | lower right watermark Aspose.Cells | keep original watermark size | PdfSaveOptions watermark example | C# add image watermark to PDF | Aspose.Cells export PDF with logo
// Common Searches: how to add an image watermark to a PDF with Aspose.Cells | Aspose.Cells place watermark bottom right | keep original size of watermark in Aspose.Cells PDF | C# RenderingWatermark alignment options | Aspose.Cells PDF export watermark opacity
// Developer Intent: Insert an image watermark into the PDF generated from an Aspose.Cells workbook, positioned at the lower‑right corner without scaling.
// Use Cases: Brand reports with a company logo placed at the bottom‑right of each PDF page. | Add a confidential stamp image to exported PDFs while preserving its exact size. | Overlay a seal image as a foreground watermark on multi‑sheet PDF exports.
// AI Prompts: Generate C# code using Aspose.Cells to add a PNG watermark to a PDF, aligned bottom‑right, original size, fully opaque. | Explain how to change the watermark to a background element and rotate it 45 degrees in the given code. | Show how to apply different image watermarks to individual worksheets before saving them as a single PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing;

namespace AsposeCellsExamples
{
    // Creates an in‑memory workbook, loads a PNG image, configures a RenderingWatermark with right and bottom alignment, 100 % scale, no rotation, full opacity, and applies it via PdfSaveOptions to produce a PDF where the watermark appears in the lower‑right corner at its original dimensions.
    public class ImageWatermarkLowerRight
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (in‑memory)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample data for watermark demonstration.");

                // Load the watermark image bytes (replace with your actual image path)
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
                    // Position at lower‑right corner
                    HAlignment = TextAlignmentType.Right,
                    VAlignment = TextAlignmentType.Bottom,

                    // Keep original size (no scaling)
                    ScaleToPagePercent = 100,

                    // No rotation, fully opaque
                    Rotation = 0f,
                    Opacity = 1f,

                    // Place on top of page contents (optional)
                    IsBackground = false
                };

                // Configure PDF save options with the watermark
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Watermark = watermark
                };

                // Save the workbook as PDF with the watermark applied
                string outputPath = "WatermarkedOutput.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ImageWatermarkLowerRight.Run();
        }
    }
}
