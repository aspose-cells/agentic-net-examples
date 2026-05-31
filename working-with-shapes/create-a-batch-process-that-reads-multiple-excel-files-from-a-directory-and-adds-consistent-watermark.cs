using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace BatchWatermark
{
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the source Excel files
            string sourceDirectory = @"C:\InputExcelFiles";

            // Directory where the watermarked PDFs will be saved
            string outputDirectory = @"C:\WatermarkedPdfs";

            // Verify source directory exists
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory not found: {sourceDirectory}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Define a consistent watermark font
            RenderingFont watermarkFont = new RenderingFont("Arial", 48)
            {
                Bold = true,
                Color = System.Drawing.Color.FromArgb(128, 0, 0, 255) // Semi‑transparent blue
            };

            // Create a single watermark instance that will be reused for all files
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                Rotation = 45f,
                Opacity = 0.3f,
                ScaleToPagePercent = 75,
                IsBackground = true
            };

            // Process each Excel file in the source directory
            foreach (string excelPath in Directory.GetFiles(sourceDirectory, "*.xlsx"))
            {
                try
                {
                    // Verify the Excel file exists before loading
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: {excelPath}");
                        continue;
                    }

                    // Load the workbook from file
                    Workbook workbook = new Workbook(excelPath);

                    // Prepare PDF save options with the watermark
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        Watermark = watermark
                    };

                    // Determine output PDF file name
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                    string pdfPath = Path.Combine(outputDirectory, fileNameWithoutExt + "_Watermarked.pdf");

                    // Save the workbook as PDF with the watermark applied
                    workbook.Save(pdfPath, pdfOptions);

                    Console.WriteLine($"Watermarked PDF created: {pdfPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
                }
            }
        }
    }
}