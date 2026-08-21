// Title: Export macro‑enabled Excel (.xlsm) to PDF with VBA retained using Aspose.Cells for .NET
// Description: Demonstrates how to load a .xlsm workbook with LoadOptions (LoadFormat.Xlsx) so the VBA macros stay intact, then save the file as a PDF using PdfSaveOptions with ExportDocumentStructure enabled, leaving the original macro‑enabled workbook unchanged.
// Keywords: Aspose.Cells C# export PDF | macro enabled Excel to PDF | retain VBA macros Aspose | LoadOptions Xlsx macro workbook | PdfSaveOptions ExportDocumentStructure | convert .xlsm to PDF .NET | preserve macros Aspose.Cells
// Common Searches: Aspose.Cells keep VBA macros when converting .xlsm to PDF | Load macro‑enabled workbook without stripping macros C# | Export Excel macro workbook to PDF preserving document structure | How to save .xlsm as PDF with macros using Aspose.Cells | PdfSaveOptions settings for macro‑enabled Excel files
// Developer Intent: Load a macro‑enabled Excel file, ensure the VBA code remains available, and generate a PDF version of the workbook.
// Use Cases: Load an .xlsm file with LoadOptions(LoadFormat.Xlsx) to keep workbook.HasMacro true. | Check workbook.HasMacro before exporting to confirm macros are present. | Save the workbook to PDF using PdfSaveOptions where ExportDocumentStructure = true, while the source file retains its macros.
// AI Prompts: Write C# code that opens a .xlsm file with Aspose.Cells, preserves its VBA macros, and exports it to PDF with document structure enabled. | Explain how LoadOptions and PdfSaveOptions work together to retain macros during PDF conversion in Aspose.Cells. | Provide troubleshooting steps if workbook.HasMacro returns false after loading a macro‑enabled workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load a .xlsm workbook with LoadOptions (LoadFormat.Xlsx) so the VBA macros stay intact, then save the file as a PDF using PdfSaveOptions with ExportDocumentStructure enabled, leaving the original macro‑enabled workbook unchanged.
class RetainMacroExportPdf
{
    static void Main()
    {
        try
        {
            // Path to the macro‑enabled Excel file
            string inputPath = "macroWorkbook.xlsm";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load options – use Xlsx format (covers .xlsm) to retain macros
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified load options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Confirm that macros are present
            Console.WriteLine("Workbook contains macros: " + workbook.HasMacro);

            // Configure PDF save options (e.g., export document structure)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Export the workbook to PDF while keeping the macros in the original workbook
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
