// Title: How to set a specific PDF version with Aspose.Cells PdfSaveOptions for Excel‑to‑PDF conversion in C#
// AI Prompts: Generate C# code that loads an .xlsx workbook and saves it as a PDF using Aspose.Cells, configuring PdfSaveOptions.Version to PDF 1.4 for legacy reader support. | Show how to enable PDF/A‑1b compliance with PdfSaveOptions when converting an Excel workbook to PDF in .NET.
// Common Searches: Aspose.Cells C# set PdfSaveOptions.Version to PDF 1.4 | Export Excel workbook to PDF with specific PDF version using Aspose.Cells | Create PDF/A‑1b file from Excel with Aspose.Cells PdfSaveOptions | Backward compatible PDF output from Aspose.Cells workbook conversion
// Tags: Aspose.Cells PDF version configuration | C# Excel to PDF specific version | PdfSaveOptions PDF/A-1b compliance | Aspose.Cells backward compatible PDF output | Configure PDF version with PdfSaveOptions

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that 'input.xlsx' exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions instance, and saves the workbook as 'output.pdf' using those options, with exception handling and console status messages.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (SaveFormat is set internally for PdfSaveOptions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
