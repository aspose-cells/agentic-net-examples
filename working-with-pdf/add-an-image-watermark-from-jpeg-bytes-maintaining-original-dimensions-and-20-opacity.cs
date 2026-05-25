using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class AddImageWatermark
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample content with image watermark");

            // Load JPEG image bytes if the file exists
            string imagePath = "watermark.jpg";
            RenderingWatermark watermark = null;
            if (File.Exists(imagePath))
            {
                byte[] imageBytes = File.ReadAllBytes(imagePath);
                watermark = new RenderingWatermark(imageBytes)
                {
                    Opacity = 0.2f,                     // 20% opacity
                    ScaleToPagePercent = 100,           // keep original dimensions
                    IsBackground = true,                // place behind page contents
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center
                };
            }
            else
            {
                Console.WriteLine($"Warning: Image file '{imagePath}' not found. PDF will be saved without watermark.");
            }

            // Configure PDF save options and attach watermark if available
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            if (watermark != null)
                pdfOptions.Watermark = watermark;

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