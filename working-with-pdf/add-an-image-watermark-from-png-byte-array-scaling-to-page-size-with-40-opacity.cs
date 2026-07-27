using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample content
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data");

        // Load PNG image into a byte array (ensure the file exists at the specified path)
        string imagePath = "watermark.png";
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Image file not found: {imagePath}");
            return;
        }

        byte[] imageData = File.ReadAllBytes(imagePath);

        // Create an image watermark from the byte array
        RenderingWatermark watermark = new RenderingWatermark(imageData);
        watermark.Opacity = 0.4f;                 // 40% opacity
        watermark.ScaleToPagePercent = 100;      // Scale to full page size
        watermark.IsBackground = true;           // Place behind page contents
        watermark.HAlignment = TextAlignmentType.Center;
        watermark.VAlignment = TextAlignmentType.Center;

        // Apply the watermark via PDF save options
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Watermark = watermark;

        // Save the workbook as a PDF with the watermark applied
        workbook.Save("output_watermark.pdf", saveOptions);
    }
}

// Author: Aspose.Cells .NET example (image watermark from byte array)