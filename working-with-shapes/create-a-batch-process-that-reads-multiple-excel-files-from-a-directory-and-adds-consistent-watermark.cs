// Title: Batch add a consistent watermark to Excel files and convert them to PDF with Aspose.Cells for .NET
// Description: A C# console app that scans a folder for .xlsx workbooks, loads each with Aspose.Cells, applies a single semi‑transparent diagonal "CONFIDENTIAL" RenderingWatermark, and saves the result as a PDF in a target directory. The solution handles missing files, creates output folders automatically, and logs processing status.
// Keywords: Aspose.Cells batch watermark | C# add watermark to Excel PDF | RenderingWatermark example | convert multiple .xlsx to PDF | semi transparent diagonal watermark | folder processing Aspose.Cells | PDFSaveOptions watermark | automated Excel to PDF conversion
// Common Searches: how to batch watermark Excel files with Aspose.Cells | C# convert folder of .xlsx to PDF with watermark | Aspose.Cells RenderingWatermark for multiple workbooks | apply same diagonal watermark to many Excel PDFs | automate Excel to PDF conversion with watermark .NET
// Developer Intent: The developer needs to process a directory of Excel workbooks, convert each to PDF, and apply an identical semi‑transparent diagonal watermark in a single pass.
// Use Cases: Produce confidential PDFs from a batch of financial spreadsheets before external sharing. | Create branded marketing PDFs from Excel templates with a company logo watermark. | Automate legal document distribution by adding a "CONFIDENTIAL" watermark to all Excel‑derived PDFs.
// AI Prompts: Generate a C# script that reads all .xls and .xlsx files in a folder, adds a custom text watermark using Aspose.Cells RenderingWatermark, and saves each as a PDF. | Explain how to vary watermark opacity and rotation based on workbook metadata when batch processing with Aspose.Cells. | Show how to write a CSV log of processed files, including success/failure status, while applying watermarks to Excel workbooks.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace BatchWatermark
{
    // A C# console app that scans a folder for .xlsx workbooks, loads each with Aspose.Cells, applies a single semi‑transparent diagonal "CONFIDENTIAL" RenderingWatermark, and saves the result as a PDF in a target directory. The solution handles missing files, creates output folders automatically, and logs processing status.
    class Program
    {
        static void Main(string[] args)
        {
            // Input directory containing Excel files
            string inputDir = @"C:\InputExcelFiles";
            // Output directory for watermarked PDFs
            string outputDir = @"C:\WatermarkedPdfs";

            // Verify input directory exists
            if (!Directory.Exists(inputDir))
            {
                Console.WriteLine($"Input directory not found: {inputDir}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputDir);

            // Define a consistent watermark font
            RenderingFont watermarkFont = new RenderingFont("Calibri", 68)
            {
                Bold = true,
                Italic = true,
                Color = Color.FromArgb(128, 0, 0, 255) // Semi‑transparent blue
            };

            // Create a single watermark instance to be reused
            RenderingWatermark watermark = new RenderingWatermark("CONFIDENTIAL", watermarkFont)
            {
                HAlignment = TextAlignmentType.Center,
                VAlignment = TextAlignmentType.Center,
                Rotation = 45f,
                Opacity = 0.3f,
                ScaleToPagePercent = 75,
                IsBackground = true
            };

            try
            {
                // Iterate over all .xlsx files in the input directory
                foreach (string excelPath in Directory.GetFiles(inputDir, "*.xlsx"))
                {
                    // Verify the file exists before loading
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found (skipped): {excelPath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook from file
                        Workbook workbook = new Workbook(excelPath);

                        // Prepare PDF save options with the watermark
                        PdfSaveOptions pdfOptions = new PdfSaveOptions
                        {
                            Watermark = watermark
                        };

                        // Build output PDF file name (same base name as Excel file)
                        string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                        string pdfPath = Path.Combine(outputDir, pdfFileName);

                        // Save the workbook as PDF with the watermark applied
                        workbook.Save(pdfPath, pdfOptions);

                        Console.WriteLine($"Processed '{excelPath}' -> '{pdfPath}'");
                    }
                    catch (Exception exFile)
                    {
                        Console.WriteLine($"Error processing file '{excelPath}': {exFile.Message}");
                    }
                }

                Console.WriteLine("Batch watermarking completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
