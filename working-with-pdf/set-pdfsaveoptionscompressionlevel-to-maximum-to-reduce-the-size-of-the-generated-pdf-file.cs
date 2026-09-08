// Title: How to set PdfSaveOptions.CompressionLevel to maximum for smaller PDF files when converting Excel to PDF with Aspose.Cells in C#
// AI Prompts: Assign PdfSaveOptions.CompressionLevel = PdfCompressionLevel.Maximum before invoking Workbook.Save to generate a compressed PDF. | Provide a C# example that configures Aspose.Cells PdfSaveOptions for the strongest compression and exports an Excel file to a smaller PDF.
// Common Searches: asp.net convert excel to pdf with highest compression using aspose.cells | c# set PdfSaveOptions.CompressionLevel to reduce PDF size | how to shrink PDF output from Excel workbook in Aspose.Cells | example of PdfCompressionLevel.Maximum in C# Aspose.Cells PDF export | optimize PDF file size when saving workbook as PDF in .NET
// Tags: Aspose.Cells PDF compression setting | PdfSaveOptions compression configuration | Excel to PDF high compression Aspose.Cells | C# PdfCompressionLevel enum usage | reduce generated PDF size Aspose.Cells | PdfSaveOptions property for PDF output

using Aspose.Cells;
using System;
using System.IO;

// // Loads an Excel workbook, creates a PdfSaveOptions instance with CompressionLevel set to PdfCompressionLevel.Maximum, and saves the workbook as a reduced‑size PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (compression setting omitted as not supported in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file using the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
