// Title: Render Office Add‑Ins and convert an XLSX workbook to PDF with default scaling using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, ensures Office Add‑In content is rendered, and saves the workbook as a PDF using Aspose.Cells with default PdfSaveOptions. | Demonstrate how to use Aspose.Cells PdfSaveOptions in C# to export an Excel workbook to PDF while preserving the default scaling and any embedded Office Add‑Ins.
// Common Searches: asp.net convert xlsx to pdf preserving office add‑ins with Aspose.Cells default scaling | c# Aspose.Cells render office add‑ins when saving workbook as PDF | how to use PdfSaveOptions for default scaling in Excel to PDF conversion | example of converting Excel workbook to PDF with Aspose.Cells while keeping add‑in rendering | Aspose.Cells PDF export default settings C# code sample
// Tags: Aspose.Cells PDF conversion default scaling | C# render Office Add‑Ins with Aspose.Cells | PdfSaveOptions Excel to PDF example | Excel workbook PDF export Aspose.Cells | Office Add‑In rendering in PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks for the input XLSX file, loads it into an Aspose.Cells Workbook, applies default PdfSaveOptions (which preserve Office Add‑In rendering and default scaling), and saves the workbook as a PDF, handling any errors that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default settings are sufficient)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
