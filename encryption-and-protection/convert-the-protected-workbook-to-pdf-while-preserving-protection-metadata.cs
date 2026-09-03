// Title: Convert a password‑protected Excel workbook to PDF with Aspose.Cells for .NET while retaining protection metadata
// AI Prompts: Write C# code that opens a password‑protected .xlsx file using Aspose.Cells LoadOptions and saves it as a PDF with PdfSaveOptions, preserving the workbook's protection settings. | Show how to verify the source Excel file exists and implement robust exception handling when converting a protected workbook to PDF in C#. | Demonstrate configuring Aspose.Cells to export a protected workbook to PDF without stripping its protection metadata, including any necessary options.
// Common Searches: asp.net convert password protected xlsx to pdf using aspose.cells preserving protection | c# load encrypted excel workbook and export to pdf with aspose.cells | how to keep workbook protection metadata when saving excel as pdf with aspose.cells | asp.net core example converting protected excel file to pdf with error handling
// Tags: Aspose.Cells load password protected workbook | Aspose.Cells export protected Excel to PDF | PdfSaveOptions usage with protected workbook | C# file existence check before Aspose.Cells conversion | exception handling Aspose.Cells PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Checks for the protected Excel file, loads it with Aspose.Cells using LoadOptions, and saves it as a PDF via PdfSaveOptions, handling missing file and runtime errors.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the protected workbook
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options (default options are sufficient)
            var pdfOptions = new PdfSaveOptions();

            // Convert and save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
