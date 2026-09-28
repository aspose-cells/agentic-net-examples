// Title: Convert an XLSX workbook to PDF in C# with Aspose.Cells while retaining Office Add‑In interactive controls
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, configures PdfSaveOptions for continuous pages, and saves the workbook as a PDF while preserving embedded form fields and OLE objects. | Show how to add file‑existence verification and robust exception handling to an Aspose.Cells Excel‑to‑PDF conversion that keeps Office Add‑In controls intact. | Explain how to set PdfSaveOptions.OnePagePerSheet to false in Aspose.Cells and ensure that interactive fields remain functional in the resulting PDF.
// Common Searches: asp.net core convert excel to pdf preserving form fields using aspose.cells | c# aspose.cells keep office add‑in controls when saving workbook as pdf | how to disable onepagepersheet in aspose.cells pdf conversion | preserve ole objects in pdf generated from xlsx with aspose.cells | error handling for missing excel file during aspose.cells pdf export
// Tags: Aspose.Cells PDF export retaining interactive fields | PdfSaveOptions continuous page layout | OLE object retention in Aspose.Cells PDF output | Render embedded Office Add‑In UI in PDF | Workbook load error handling with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook, sets PdfSaveOptions with OnePagePerSheet = false to produce continuous pages, relies on Aspose.Cells to retain form fields, OLE objects, and Office Add‑In controls, saves the result as a PDF, and includes comprehensive error handling.
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

            // Load the source XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep each worksheet on its own page layout (set to false to allow continuous pages)
                OnePagePerSheet = false
                // Note: Aspose.Cells automatically preserves form fields and OLE objects where possible.
            };

            // Convert and save the workbook to PDF using the configured options
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
