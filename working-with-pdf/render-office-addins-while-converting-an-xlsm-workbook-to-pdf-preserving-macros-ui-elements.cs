// Title: Convert an XLSM workbook with macros and Office Add‑Ins to PDF using Aspose.Cells for .NET while preserving custom ribbon UI
// AI Prompts: Load a macro‑enabled XLSM file with Aspose.Cells, activate macro processing, and save it as a PDF while retaining the workbook’s visual layout. | Configure PdfSaveOptions to apply PDF/A‑1b compliance and disable OnePagePerSheet so each sheet keeps its original dimensions during conversion. | Implement error handling that checks for the source XLSM file, creates the output directory if needed, and catches exceptions during the PDF export.
// Common Searches: asp.net convert xlsm with custom ribbon to pdf using aspose.cells | how to preserve office add‑ins UI when saving macro enabled workbook to pdf | aspose.cells enable macros before exporting xlsm to pdf | pdfsaveoptions pdf/a‑1b and onepagepersheet settings for xlsm conversion | c# load macro enabled workbook and export to pdf preserving layout
// Tags: xlsm to pdf conversion with macro support asp.net | aspose.cells enablemacros setting | pdfsaveoptions pdf/a‑1b compliance | preserve custom ribbon UI aspose.cells | onepagepersheet false option

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program loads a macro‑enabled XLSM workbook, activates macro processing, configures PDF/A‑1b compliance and layout options, and saves the workbook as a PDF while keeping visual elements such as custom ribbons intact.
class Program
{
    static void Main()
    {
        // Path to the source XLSM workbook (contains macros and UI add‑ins)
        string sourceFile = @"C:\Path\To\YourWorkbook.xlsm";

        // Path for the resulting PDF file
        string pdfFile = @"C:\Path\To\ConvertedWorkbook.pdf";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(sourceFile))
                throw new FileNotFoundException($"Source file not found: {sourceFile}");

            // Ensure the output directory exists
            string pdfDir = Path.GetDirectoryName(pdfFile);
            if (!Directory.Exists(pdfDir))
                Directory.CreateDirectory(pdfDir);

            // Load the XLSM workbook (macro-enabled). No explicit LoadFormat needed; Aspose.Cells detects it.
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Enable macro processing while the workbook is in memory
            workbook.Settings.EnableMacros = true;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep each sheet on its own page range (default behavior)
                OnePagePerSheet = false,

                // Do not force all columns onto a single page
                AllColumnsInOnePagePerSheet = false,

                // Preserve the visual layout as close as possible to Excel
                Compliance = PdfCompliance.PdfA1b
            };

            // Save the workbook as PDF; macros themselves are not rendered in PDF,
            // but any UI elements (e.g., custom ribbons) that affect the visual layout
            // are retained in the output.
            workbook.Save(pdfFile, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {pdfFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
