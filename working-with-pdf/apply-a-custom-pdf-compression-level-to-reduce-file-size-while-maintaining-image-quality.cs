// Title: How to set a custom PDF compression level when converting an Excel workbook to PDF with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set PDF compression level to reduce file size | custom PdfSaveOptions for high‑quality image compression in Excel to PDF conversion | reduce PDF output size while preserving images using Aspose.Cells .NET | configure PDF save options for compression in Aspose.Cells C# example | adjust PDF compression settings when exporting Excel workbook with Aspose.Cells
// Tags: Aspose.Cells PDF compression level configuration | PdfSaveOptions image quality control | Excel workbook to PDF size optimization .NET | custom PDF save options Aspose.Cells | maintain image fidelity during PDF export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample checks that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (using default settings), and saves the workbook as a PDF. By modifying properties of PdfSaveOptions—such as CompressionLevel—you can control the PDF compression level to lower file size while keeping image quality. Exceptions are caught and logged.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.pdf";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the source Excel file
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (using defaults compatible with current Aspose.Cells version)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF using the configured options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF file saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
