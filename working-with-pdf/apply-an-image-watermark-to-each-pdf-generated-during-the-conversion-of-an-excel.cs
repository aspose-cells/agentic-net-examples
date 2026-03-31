using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Path to the image that will be used as a watermark
            string watermarkImagePath = "watermark.png";

            // Path for the resulting PDF file
            string pdfOutputPath = "output.pdf";

            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(excelPath);

            // Read the watermark image into a byte array
            if (!File.Exists(watermarkImagePath))
            {
                Console.WriteLine($"Watermark image not found: {watermarkImagePath}");
                return;
            }
            byte[] imageData = File.ReadAllBytes(watermarkImagePath);

            // Create an image‑based watermark
            RenderingWatermark watermark = new RenderingWatermark(imageData)
            {
                // Example visual settings – adjust as needed
                Rotation = 45f,                 // Rotate 45 degrees
                Opacity = 0.5f,                 // 50% transparent
                ScaleToPagePercent = 50,        // Scale to 50% of the page size
                IsBackground = true,            // Place behind page contents
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center
            };

            // Configure PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the entire workbook as a PDF with the watermark applied
            workbook.Save(pdfOutputPath, pdfOptions);

            Console.WriteLine($"Workbook converted to PDF with image watermark: {pdfOutputPath}");
        }
    }
}