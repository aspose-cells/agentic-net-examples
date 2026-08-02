// Title: Add JPEG Image Watermark to PDF from Aspose.Cells (20% Opacity, Original Size)
// Description: C# example that creates an empty Workbook, loads a JPEG file as a byte array, builds a RenderingWatermark with 20% opacity, keeps the original image dimensions (ScaleToPagePercent = 100) and places it behind the worksheet content. The watermark is attached to PdfSaveOptions and the workbook is saved as a PDF.
// Keywords: Aspose.Cells PDF watermark C# | RenderingWatermark JPEG bytes | image watermark opacity 0.2 | keep original watermark size | background watermark Aspose.Cells | PdfSaveOptions watermark | C# add image watermark to PDF
// Common Searches: Aspose.Cells add JPEG watermark to PDF | C# set watermark opacity to 20% in PDF | keep original image size when watermarking PDF with Aspose.Cells | background image watermark Aspose.Cells PDF export | RenderingWatermark example C#
// Developer Intent: Apply a JPEG image as a background watermark to a PDF generated from an Aspose.Cells workbook, preserving the image’s original dimensions and using 20% opacity.
// Use Cases: Brand a PDF report with a faint company logo behind the data. | Create confidential Excel‑to‑PDF exports that include a light watermark to deter unauthorized copying. | Add a custom background image to every page of PDFs produced from Excel workbooks.
// AI Prompts: Show how to change the watermark to a PNG file with 30% opacity. | Explain how to apply different watermarks to individual worksheets when exporting a single PDF. | Provide code to center the watermark on each page and rotate it 45 degrees.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    // C# example that creates an empty Workbook, loads a JPEG file as a byte array, builds a RenderingWatermark with 20% opacity, keeps the original image dimensions (ScaleToPagePercent = 100) and places it behind the worksheet content. The watermark is attached to PdfSaveOptions and the workbook is saved as a PDF.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (empty workbook with a default worksheet)
                Workbook workbook = new Workbook();

                const string watermarkPath = "watermark.jpg";

                // Verify that the watermark image file exists
                if (!File.Exists(watermarkPath))
                {
                    Console.WriteLine($"Error: Watermark image file '{watermarkPath}' not found.");
                    return;
                }

                // Load JPEG image bytes
                byte[] jpegBytes = File.ReadAllBytes(watermarkPath);

                // Create an image‑based watermark using the JPEG bytes
                RenderingWatermark watermark = new RenderingWatermark(jpegBytes)
                {
                    // Set opacity to 20% (0.2) and keep the original image size
                    Opacity = 0.2f,               // 20% opacity
                    ScaleToPagePercent = 100,    // 100% keeps original dimensions
                    // Place the watermark behind the worksheet content
                    IsBackground = true
                };

                // Configure PDF save options with the watermark
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Watermark = watermark
                };

                // Save the workbook as a PDF with the image watermark applied
                const string outputPath = "WatermarkedOutput.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
