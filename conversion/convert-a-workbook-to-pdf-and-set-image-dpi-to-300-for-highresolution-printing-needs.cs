// Title: Export Aspose.Cells Workbook to PDF with 300 DPI Images in C#
// Description: Shows how to set the global DPI to 300, use PdfSaveOptions to resample images at 300 PPI with full JPEG quality, and save the workbook as a print‑ready PDF.
// Keywords: Aspose.Cells PDF export C# | 300 DPI PDF Aspose.Cells | PdfSaveOptions SetImageResample | high resolution PDF from workbook | C# image DPI conversion Aspose
// Common Searches: Aspose.Cells export PDF 300 DPI | C# set image resample when saving to PDF | increase PDF image quality Aspose.Cells | global DPI setting Aspose.Cells workbook | PdfSaveOptions image DPI C# example
// Developer Intent: Generate a PDF where every embedded image is rendered at 300 DPI for professional printing.
// Use Cases: Print‑ready financial reports with crisp charts and logos. | High‑quality invoices where company branding must stay sharp. | Batch conversion of multiple workbooks to PDFs that meet publishing standards.
// AI Prompts: Modify the sample to use 80 % JPEG quality while keeping 300 DPI image resampling. | Provide a C# snippet that merges all worksheets into a single high‑DPI PDF. | Explain how to combine font embedding with image DPI settings in PdfSaveOptions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;   // Required for PdfSaveOptions

namespace AsposeCellsPdfConversion
{
    // Shows how to set the global DPI to 300, use PdfSaveOptions to resample images at 300 PPI with full JPEG quality, and save the workbook as a print‑ready PDF.
    public class ConvertToPdfHighDpi
    {
        public static void Run()
        {
            try
            {
                // Set global DPI for rendering operations.
                CellsHelper.DPI = 300;

                // Create a simple workbook with sample data.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("High‑Resolution PDF Export");
                sheet.Cells["A2"].PutValue(DateTime.Now);
                sheet.Cells["B1"].PutValue(123);
                sheet.Cells["B2"].PutValue(456);

                // Configure PDF save options: resample images at 300 PPI, 100% JPEG quality.
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                pdfOptions.SetImageResample(300, 100);

                // Define output path and ensure its directory exists.
                string outputPath = "HighResolutionOutput.pdf";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF file using the configured options.
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as high‑resolution PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error during PDF conversion: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertToPdfHighDpi.Run();
        }
    }
}
