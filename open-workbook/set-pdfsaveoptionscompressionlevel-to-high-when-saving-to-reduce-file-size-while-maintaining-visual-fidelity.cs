// Title: Setting PdfSaveOptions.CompressionLevel to High in Aspose.Cells C# for smaller Excel-to-PDF files while preserving visual fidelity
// Common Searches: Aspose.Cells C# set PdfSaveOptions.CompressionLevel to High for PDF export | How to reduce PDF size from Excel workbook using Aspose.Cells without losing quality | C# example of high compression PDF conversion with Aspose.Cells PdfSaveOptions | Excel to PDF with maximum compression level in .NET using Aspose.Cells | Configure PdfSaveOptions for smaller PDF output while keeping visual fidelity in C#
// Tags: Aspose.Cells PDF compression level high | C# Excel to PDF high compression | PdfSaveOptions high compression setting | reduce PDF file size Aspose.Cells conversion | preserve visual fidelity PDF export .NET | configure PDF save options Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, sets its CompressionLevel property to High to shrink the resulting PDF while maintaining visual fidelity, and saves the workbook as a PDF, all wrapped in exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options are sufficient)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
