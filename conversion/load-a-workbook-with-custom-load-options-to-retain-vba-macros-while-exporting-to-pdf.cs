// Title: Convert a macro‑enabled XLSM workbook to PDF while preserving VBA macros using Aspose.Cells for .NET
// AI Prompts: Load an .xlsm file with Aspose.Cells, verify workbook.HasMacro, set PdfSaveOptions.ExportDocumentStructure = true, and save as PDF. | Demonstrate how to retain VBA macros when converting a macro‑enabled Excel workbook to PDF in C# using Aspose.Cells. | Show code to check file existence, load the workbook, and export to PDF with document structure while keeping macros.
// Common Searches: how to keep VBA macros when converting xlsm to pdf using Aspose.Cells .NET | export macro-enabled Excel workbook to PDF with document structure using Aspose.Cells | C# example loading .xlsm and saving as PDF while preserving macros | using PdfSaveOptions ExportDocumentStructure in Aspose.Cells | determine if loaded workbook has macros with Aspose.Cells
// Tags: xlsm to pdf conversion with macro preservation | Aspose.Cells PdfSaveOptions ExportDocumentStructure | load macro-enabled workbook C# Aspose.Cells | check VBA macros Aspose.Cells Workbook.HasMacro | export macro-enabled Excel to PDF .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example verifies the presence of an input .xlsm file, loads it with Aspose.Cells (which automatically retains VBA macros), confirms macro existence via workbook.HasMacro, configures PdfSaveOptions to export document structure, and saves the workbook as a PDF.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsm";
                string outputPath = "output.pdf";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook; macros are retained automatically
                Workbook workbook = new Workbook(inputPath);

                // Optional: confirm that macros are present
                Console.WriteLine("Workbook contains macros: " + workbook.HasMacro);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true
                };

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook saved to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
