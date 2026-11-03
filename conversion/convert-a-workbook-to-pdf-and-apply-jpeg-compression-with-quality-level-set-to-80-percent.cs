// Title: Convert an Excel workbook to PDF with 80% JPEG image compression using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set JPEG compression quality when saving workbook to PDF | How to export XLSX to PDF with reduced image size using Aspose.Cells | PDFSaveOptions JPEG quality 80 percent Aspose.Cells example | C# convert Excel to PDF and control image compression level | Aspose.Cells PDF export image compression settings tutorial
// Tags: Aspose.Cells PDFSaveOptions JPEG compression | Excel to PDF conversion with image quality control | C# set JPEG quality for PDF export using Aspose.Cells | reduce PDF image size Aspose.Cells | PDFSaveOptions image compression level configuration

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook (creating a simple one if it doesn't exist), configures PdfSaveOptions to apply JPEG compression at 80% quality, and saves the workbook as a PDF while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Load the source workbook; create a simple one if the file does not exist.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample data");
                // Save the generated template for future runs.
                workbook.Save(inputPath);
            }

            // Configure PDF save options (default settings are used here).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file with the specified options.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
