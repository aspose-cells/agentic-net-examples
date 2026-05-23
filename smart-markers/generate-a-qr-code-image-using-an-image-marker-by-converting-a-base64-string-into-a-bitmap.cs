using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsQrCodeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Base64 string for a 1x1 pixel PNG image (valid Base64)
                const string qrBase64 =
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X0ZcAAAAASUVORK5CYII=";

                // Convert Base64 string to byte array
                byte[] imageBytes = Convert.FromBase64String(qrBase64);

                // Add the QR code image to the worksheet using a memory stream
                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    // Row 2, Column 2 (zero‑based indices)
                    sheet.Pictures.Add(1, 1, ms);
                }

                // Save the workbook to a file
                string outputPath = "QrCodeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}