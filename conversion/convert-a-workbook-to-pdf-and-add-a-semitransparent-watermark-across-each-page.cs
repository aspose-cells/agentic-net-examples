using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsWatermarkPdf
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Create a font for the watermark text
            RenderingFont font = new RenderingFont("Calibri", 68)
            {
                Bold = true,
                Italic = true,
                Color = Color.Blue
            };

            // Create a text watermark with the desired text and font
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", font)
            {
                // Center the watermark on each page
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                // Rotate for a diagonal appearance
                Rotation = 45,
                // Semi‑transparent (30% opacity)
                Opacity = 0.3f,
                // Scale relative to page size (optional)
                ScaleToPagePercent = 75,
                // Place behind the page content
                IsBackground = true
            };

            // Configure PDF save options and assign the watermark
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Watermark = watermark
            };

            // Save the workbook as a PDF with the watermark applied
            string outputPath = "output_watermarked.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF with watermark: {outputPath}");
        }
    }
}