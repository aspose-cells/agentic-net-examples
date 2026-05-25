using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfCompressionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF Compression Demo");
            sheet.Cells["A2"].PutValue("This PDF uses Flate compression and high‑quality image resampling.");

            // Optionally add an image to demonstrate image handling
            // Ensure the image file exists at the specified path
            // sheet.Pictures.Add(5, 0, "sample_image.jpg");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Use Flate compression for non‑image content (good compression ratio)
            pdfOptions.PdfCompression = PdfCompressionCore.Flate;

            // Optimize for smaller file size while keeping print quality
            pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

            // Resample images to a high PPI (220) with high JPEG quality (90%)
            // This keeps image quality high while allowing compression
            pdfOptions.SetImageResample(220, 90);

            // Save the workbook as a PDF with the specified options
            workbook.Save("CompressedHighQuality.pdf", pdfOptions);

            Console.WriteLine("PDF saved with custom compression and high image quality.");
        }
    }
}