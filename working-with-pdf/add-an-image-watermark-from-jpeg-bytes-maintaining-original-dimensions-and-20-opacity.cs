// Title: Add JPEG Image Watermark to PDF from Aspose.Cells Using Byte Array (20% Opacity)
// Description: Creates a Workbook, reads a JPEG file into a byte array, builds a RenderingWatermark with 20% opacity, keeps the original dimensions, centers the image, and applies it to PdfSaveOptions. The PDF is saved with the watermark, and the code gracefully skips the watermark if the image file is missing.
// Keywords: Aspose.Cells PDF watermark | RenderingWatermark byte array | JPEG watermark opacity | scale watermark original size | C# add image watermark | PdfSaveOptions watermark | conditional watermark Aspose.Cells
// Common Searches: Aspose.Cells add image watermark to PDF | C# RenderingWatermark from byte array | set watermark opacity 20% Aspose.Cells | preserve watermark dimensions PDF export | apply JPEG watermark only if file exists
// Developer Intent: Create a PDF from a workbook and overlay a JPEG watermark loaded from bytes, keeping its original size and 20% opacity.
// Use Cases: Brand a generated PDF report with a semi‑transparent company logo stored as a byte array. | Add a faint background image to invoices when the logo file is available. | Export spreadsheets to PDF with optional watermark based on file existence.
// AI Prompts: Generate C# code that uses Aspose.Cells to apply a PNG watermark from a MemoryStream to a PDF with 30% opacity and 50% scaling. | Describe how RenderingWatermark properties Opacity, ScaleToPagePercent, IsBackground, HAlignment, and VAlignment control the visual result of an image watermark in a PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a Workbook, reads a JPEG file into a byte array, builds a RenderingWatermark with 20% opacity, keeps the original dimensions, centers the image, and applies it to PdfSaveOptions. The PDF is saved with the watermark, and the code gracefully skips the watermark if the image file is missing.
class AddImageWatermark
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for watermark demonstration");

            RenderingWatermark watermark = null;
            string imagePath = "watermark.jpg";

            // Load JPEG image bytes if the file exists
            if (File.Exists(imagePath))
            {
                byte[] jpegBytes = File.ReadAllBytes(imagePath);
                watermark = new RenderingWatermark(jpegBytes);
                // Set watermark properties
                watermark.Opacity = 0.2f;                     // 20% opacity
                watermark.ScaleToPagePercent = 100;          // Keep original dimensions
                watermark.IsBackground = false;              // Place in front of content
                watermark.HAlignment = TextAlignmentType.Center;
                watermark.VAlignment = TextAlignmentType.Center;
            }
            else
            {
                Console.WriteLine($"Warning: Image file '{imagePath}' not found. Saving without watermark.");
            }

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Apply watermark only if it was created
            if (watermark != null)
            {
                pdfOptions.Watermark = watermark;
            }

            // Save the workbook as a PDF
            workbook.Save("WatermarkedOutput.pdf", pdfOptions);
            Console.WriteLine("PDF saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
