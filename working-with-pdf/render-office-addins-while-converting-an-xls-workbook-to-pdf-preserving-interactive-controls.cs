// Title: Convert legacy XLS workbook to PDF in C# using Aspose.Cells PdfSaveOptions
// Common Searches: C# Aspose.Cells example for converting .xls to PDF with default options | How to export an Excel 97‑2003 workbook as PDF using Aspose.Cells .NET | Aspose.Cells PdfSaveOptions usage for XLS to PDF conversion in C# | Saving legacy Excel files to PDF programmatically with Aspose.Cells
// Tags: Aspose.Cells XLS to PDF conversion C# | PdfSaveOptions Excel to PDF export | legacy Excel workbook PDF conversion .NET | C# programmatic Excel 97-2003 to PDF | file existence check Aspose.Cells conversion

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an XLS workbook with Aspose.Cells, applies default PdfSaveOptions, and saves it as a PDF file, including basic error handling for missing input files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the source XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options are sufficient for basic conversion)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
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
