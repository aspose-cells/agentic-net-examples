using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample content
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Sample data for watermark demonstration");

            // Prepare an image‑based watermark if the file exists
            RenderingWatermark watermark = null;
            string watermarkPath = "watermark.png";

            if (File.Exists(watermarkPath))
            {
                byte[] imageBytes = File.ReadAllBytes(watermarkPath);
                watermark = new RenderingWatermark(imageBytes)
                {
                    ScaleToPagePercent = 100,   // cover the whole page
                    Opacity = 0.4f,             // 40% opacity
                    IsBackground = true,        // behind worksheet content
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center
                };
            }
            else
            {
                Console.WriteLine($"Warning: Watermark image '{watermarkPath}' not found. PDF will be saved without a watermark.");
            }

            // Configure PDF save options, attaching the watermark if available
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