using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample data for PDF with image watermark.");

                // -----------------------------------------------------------------
                // 1. Prepare the image watermark (e.g., a logo)
                // -----------------------------------------------------------------
                string imagePath = "logo.png"; // Path to your image file
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Load the image bytes
                byte[] imageBytes = File.ReadAllBytes(imagePath);

                // -----------------------------------------------------------------
                // 2. Create a RenderingWatermark from the image bytes
                // -----------------------------------------------------------------
                RenderingWatermark watermark = new RenderingWatermark(imageBytes)
                {
                    // Appear above the page contents
                    IsBackground = false,

                    // Center it on each page
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center,

                    // Scale it to a reasonable size relative to the page
                    ScaleToPagePercent = 50,

                    // Slight opacity so underlying content is still readable
                    Opacity = 0.4f,

                    // No rotation needed
                    Rotation = 0
                };

                // -----------------------------------------------------------------
                // 3. Configure PDF save options with the watermark
                // -----------------------------------------------------------------
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Watermark = watermark
                };

                // -----------------------------------------------------------------
                // 4. Save the workbook as PDF with the watermark
                // -----------------------------------------------------------------
                string outputPdf = "Workbook_With_Image_Watermark.pdf";
                workbook.Save(outputPdf, pdfOptions);

                Console.WriteLine($"PDF saved successfully: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}