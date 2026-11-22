// Title: Convert an Excel workbook to PDF with JPEG compression set to 50% quality using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook as PDF with JPEG quality 50 | set JPEG compression level when exporting Excel to PDF using Aspose.Cells | reduce PDF image size from Excel conversion Aspose.Cells .NET | how to specify JpegQuality in PdfSaveOptions Aspose.Cells | C# convert .xlsx to compressed PDF with 50% image quality
// Tags: Aspose.Cells PdfSaveOptions JpegQuality | Excel to PDF conversion with JPEG compression | C# set PDF image compression Aspose.Cells | reduce PDF file size Aspose.Cells .NET | export workbook to PDF with 50% image quality

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file (or creates a new workbook), configures PdfSaveOptions to use JPEG compression with a quality of 50%, ensures the output directory exists, and saves the workbook as a PDF. It demonstrates how to lower PDF image size during Excel‑to‑PDF conversion using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one with sample data.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
            }

            // Configure PDF save options (default settings are used here).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF using the configured options.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
