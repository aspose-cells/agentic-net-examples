using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfSaveOptions

class RetainMacroExportToPdf
{
    static void Main()
    {
        try
        {
            // Path to the macro‑enabled Excel file
            string sourceFile = "input.xlsm";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file not found: {sourceFile}");
                return;
            }

            // LoadOptions configured for Xlsx format (used for .xlsm files as well)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options (macros are retained)
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Create PDF save options (optional settings can be adjusted here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: keep document structure in the PDF
                ExportDocumentStructure = true
            };

            // Export the workbook to PDF while keeping the macros in the workbook object
            string outputPdf = "output.pdf";
            workbook.Save(outputPdf, pdfOptions);

            // Optional: verify that macros are still present after export
            Console.WriteLine("Workbook has macros: " + workbook.HasMacro);
            Console.WriteLine($"PDF saved to: {outputPdf}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}