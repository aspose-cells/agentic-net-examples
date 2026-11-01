// Title: Convert an XLSX workbook to PDF with default scaling and render Office Add‑Ins using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing .xlsx file, verifies its existence, and saves it as a PDF with default scaling using Aspose.Cells PdfSaveOptions, handling any exceptions. | Show how to configure Aspose.Cells to render embedded Office Add‑Ins when exporting an Excel workbook to PDF in a .NET console application. | Create a minimal C# console program that converts an Excel workbook to PDF, includes basic file‑not‑found checking, and uses default PDF options for scaling.
// Common Searches: Aspose.Cells render Office Add‑Ins during XLSX to PDF conversion C# | C# default scaling PDF export with Aspose.Cells PdfSaveOptions | How to check if Excel file exists before converting to PDF using Aspose.Cells | Console app example for converting Excel workbook to PDF with error handling Aspose.Cells | Export Excel workbook to PDF preserving embedded objects Aspose.Cells .NET
// Tags: Aspose.Cells XLSX to PDF conversion with default scaling | render Office Add‑Ins in PDF export Aspose.Cells | PdfSaveOptions default settings .NET | file existence validation before Aspose.Cells conversion | exception handling for Excel to PDF export C#

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the input XLSX file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object with default settings (including default scaling and rendering of Office Add‑Ins), and saves the workbook as a PDF while catching and reporting any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options are sufficient for basic conversion)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Convert and save the workbook to PDF using the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
