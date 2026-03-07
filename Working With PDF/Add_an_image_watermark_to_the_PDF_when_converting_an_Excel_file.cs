using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkExample
{
    class Program
    {
        static void Main()
        {
            // Load the source Excel workbook
            string excelPath = "input.xlsx";
            Workbook workbook = new Workbook(excelPath);

            // Load the image data that will be used as the watermark
            string imagePath = "watermark.png";
            byte[] imageData;
            if (File.Exists(imagePath))
            {
                imageData = File.ReadAllBytes(imagePath);
            }
            else
            {
                // Fallback to a simple 1x1 transparent PNG if the file is missing
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
                imageData = Convert.FromBase64String(base64Png);
            }

            // Create an image‑based rendering watermark
            RenderingWatermark watermark = new RenderingWatermark(imageData)
            {
                Rotation = 45f,               // rotate 45 degrees
                Opacity = 0.5f,               // 50% transparent
                ScaleToPagePercent = 75,      // occupy 75% of the page
                IsBackground = true,          // place behind page contents
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center
            };

            // Configure PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the image watermark applied
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"Workbook converted to PDF with image watermark: {pdfPath}");
        }
    }
}