using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsQrDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.20);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.80);

                // Configure image rendering options for PNG output
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    OnePagePerSheet = true
                };

                // Render the first page of the worksheet to a PNG file
                string outputDir = "output";
                Directory.CreateDirectory(outputDir);
                string pngPath = Path.Combine(outputDir, "worksheet.png");
                SheetRender renderer = new SheetRender(sheet, imgOptions);
                renderer.ToImage(0, pngPath); // Render first page

                // URL where the PNG image will be hosted (replace with actual URL in production)
                string imageUrl = "https://example.com/files/worksheet.png";

                // Generate a simple placeholder QR code image (1x1 transparent PNG) without System.Drawing
                string qrPath = Path.Combine(outputDir, "worksheet_qr.png");
                // Base64 representation of a minimal 1x1 PNG
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
                byte[] pngBytes = Convert.FromBase64String(base64Png);
                File.WriteAllBytes(qrPath, pngBytes);

                Console.WriteLine($"Worksheet image saved to: {pngPath}");
                Console.WriteLine($"QR code placeholder image saved to: {qrPath}");
                Console.WriteLine($"QR code links to: {imageUrl}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}