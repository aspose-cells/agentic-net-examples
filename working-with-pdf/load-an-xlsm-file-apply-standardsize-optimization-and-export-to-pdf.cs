// Title: Convert a macro‑enabled XLSM workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# code to export an XLSM file to PDF with Aspose.Cells | How to use PdfSaveOptions when saving a macro‑enabled workbook as PDF in .NET | Aspose.Cells example for converting macro‑enabled Excel to PDF with file existence check | Saving an XLSM workbook as PDF in C# without losing macros | Aspose.Cells .NET tutorial for XLSM to PDF conversion with default options
// Tags: XLSM to PDF conversion with Aspose.Cells | PdfSaveOptions configuration in C# | macro-enabled workbook export to PDF | Aspose.Cells workbook.Save PDF example | input file validation before Excel conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the XLSM file exists, loads it into an Aspose.Cells Workbook, applies default PdfSaveOptions, ensures the output directory is created, and saves the workbook as a PDF, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsm";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the XLSM workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options used)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook to PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
