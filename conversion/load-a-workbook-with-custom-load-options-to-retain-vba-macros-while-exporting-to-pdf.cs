// Title: Convert a macro‑enabled .xlsm workbook to PDF with Aspose.Cells for .NET
// Description: Load an .xlsm file via Aspose.Cells, confirm VBA macros with Workbook.HasMacro, apply PdfSaveOptions.ExportDocumentStructure, and export to PDF with proper error handling.
// Keywords: Aspose.Cells .NET PDF conversion | macro-enabled workbook to PDF | xlsm to PDF Aspose.Cells | preserve VBA macros Aspose | PdfSaveOptions ExportDocumentStructure | Workbook.HasMacro check | C# Excel to PDF conversion | load xlsm with macros | Aspose.Cells conversion sample | accessible PDF from Excel
// Common Searches: Aspose.Cells convert xlsm to pdf | C# preserve VBA macros when exporting Excel to PDF | PdfSaveOptions ExportDocumentStructure example | How to check if workbook has macros Aspose.Cells | Load macro-enabled workbook with Aspose.Cells .NET
// Developer Intent: Export an Excel file that contains VBA macros to PDF without altering the original macro code.
// Use Cases: Batch process macro‑enabled Excel files and generate accessible PDFs. | Validate presence of VBA code before conversion to avoid unnecessary processing. | Create PDFs with document structure for compliance (e.g., WCAG) from workbooks containing macros. | Integrate macro detection and PDF export into automated reporting pipelines.
// AI Prompts: Generate C# code that opens an .xlsm file with Aspose.Cells, checks Workbook.HasMacro, and saves it as a PDF using PdfSaveOptions.ExportDocumentStructure. | Explain the impact of ExportDocumentStructure on PDF accessibility when converting Excel files with Aspose.Cells. | Suggest robust error‑handling patterns for loading macro‑enabled workbooks and exporting them to PDF in a .NET application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an .xlsm file via Aspose.Cells, confirm VBA macros with Workbook.HasMacro, apply PdfSaveOptions.ExportDocumentStructure, and export to PDF with proper error handling.
class ExportMacroWorkbookToPdf
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsm";
            string outputPath = "output.pdf";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the macro‑enabled workbook; format is detected automatically
            Workbook workbook = new Workbook(inputPath);

            // Verify that the workbook contains macros
            Console.WriteLine("Workbook contains macros: " + workbook.HasMacro);

            // Configure PDF save options (e.g., export document structure)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
