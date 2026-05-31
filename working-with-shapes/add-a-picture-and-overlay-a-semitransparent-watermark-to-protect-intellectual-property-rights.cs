using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a picture to the worksheet if the file exists
            string picturePath = "logo.png";
            if (File.Exists(picturePath))
            {
                int pictureIndex = sheet.Pictures.Add(2, 2, picturePath);
                // Make the picture semi‑transparent (optional)
                sheet.Pictures[pictureIndex].FormatPicture.Transparency = 0.5; // 50% transparent
            }
            else
            {
                Console.WriteLine($"Warning: Picture file \"{picturePath}\" not found. Skipping picture insertion.");
            }

            // Create a rendering font for the watermark text
            RenderingFont watermarkFont = new RenderingFont("Arial", 48)
            {
                Bold = true,
                Color = Color.LightGray
            };

            // Create a text watermark and configure its appearance
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                Opacity = 0.3f,                     // semi‑transparent
                Rotation = 45,                      // diagonal
                IsBackground = true,                // placed behind content
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center
            };

            // Set PDF save options with the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as PDF with the picture and watermark applied
            workbook.Save("output.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}