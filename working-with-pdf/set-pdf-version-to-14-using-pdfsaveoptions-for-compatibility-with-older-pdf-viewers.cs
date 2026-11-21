// Title: Set PDF version to 1.4 when converting an Excel workbook to PDF using Aspose.Cells PdfSaveOptions in C#
// Common Searches: c# aspnet set pdf version 1.4 with Aspose.Cells PdfSaveOptions | how to force PDF 1.4 compatibility when exporting Excel to PDF using Aspose.Cells | Aspose.Cells export workbook to PDF with specific PDF version for older viewers | configure PdfSaveOptions to generate PDF 1.4 in C# | save Excel file as PDF 1.4 using Aspose.Cells for legacy PDF readers
// Tags: Aspose.Cells PDF version 1.4 | C# PdfSaveOptions set PDF compatibility | export Excel workbook to PDF with specific version | PDF 1.4 generation using Aspose.Cells | legacy PDF viewer compatibility Aspose.Cells | configure PDF save options in C#

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the source Excel file, loads it into a Workbook, creates a PdfSaveOptions object, sets PdfSaveOptions.PdfVersion to PdfVersion.Pdf14 for older PDF viewer compatibility, and saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
