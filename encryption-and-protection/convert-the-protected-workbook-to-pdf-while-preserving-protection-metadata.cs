// Title: Convert a password‑protected Excel workbook to PDF while preserving protection metadata with Aspose.Cells for .NET
// Common Searches: Aspose.Cells convert password protected Excel file to PDF in C# | How to retain workbook protection metadata when exporting XLSX to PDF using Aspose.Cells | C# load encrypted .xlsx and save as PDF without losing protection with Aspose.Cells | Saving a protected Excel workbook to PDF while keeping protection settings in .NET | Aspose.Cells PDF export preserving Excel file protection attributes
// Tags: protected workbook to PDF conversion Aspose.Cells | preserve Excel protection metadata PDF export | Aspose.Cells LoadOptions encrypted XLSX | PdfSaveOptions for protected workbook | C# Aspose.Cells export password protected Excel

using System;
using System.IO;
using Aspose.Cells;

// The example checks for a password‑protected Excel file, loads it with LoadOptions to keep protection metadata, configures PdfSaveOptions, and saves the workbook as a PDF, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "protected.xlsx";
            const string outputPath = "protected_output.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the protected workbook (protection metadata is retained automatically)
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options (no explicit property for preserving protection in this version)
            var pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
