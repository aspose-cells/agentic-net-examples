// Title: Export refreshed pivot tables to PDF with layout preservation using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, refreshes all pivot tables to update calculated values, configures PdfSaveOptions to keep the original document structure, and saves the result as a PDF. Includes simple file‑existence checking and exception handling.
// Keywords: Aspose.Cells export pivot table PDF | refresh pivot tables C# | PdfSaveOptions ExportDocumentStructure | preserve Excel layout PDF | Aspose.Cells .NET PDF conversion | pivot table refresh before PDF
// Common Searches: how to refresh pivot tables before exporting to pdf with aspose.cells | c# export excel pivot table to pdf preserving layout | aspose.cells PdfSaveOptions ExportDocumentStructure example | export refreshed pivot tables to pdf using .net | save workbook as pdf with updated pivot values
// Developer Intent: Update all pivot tables in an Excel file and generate a PDF that mirrors the workbook’s visual layout.
// Use Cases: Create a PDF version of a monthly financial report after the pivot calculations have been refreshed. | Automate nightly generation of PDF dashboards from Excel workbooks, ensuring the latest pivot data and exact page formatting. | Distribute sales analysis workbooks as PDFs to stakeholders while preserving the original Excel layout and refreshed pivot results.
// AI Prompts: Generate C# code with Aspose.Cells that loads an .xlsx file, refreshes every pivot table, and saves the workbook as a PDF preserving the document structure. | Show how to set PdfSaveOptions.ExportDocumentStructure to true when exporting a workbook containing pivot tables. | Provide error‑handling examples for missing input files during pivot table refresh and PDF export with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfSaveOptions

namespace AsposeCellsPivotPdfExport
{
    // Loads an Excel workbook, refreshes all pivot tables to update calculated values, configures PdfSaveOptions to keep the original document structure, and saves the result as a PDF. Includes simple file‑existence checking and exception handling.
    public class ExportRefreshedPivotTableToPdf
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "refreshed_pivot.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook that contains the pivot table
                Workbook workbook = new Workbook(inputPath);

                // Refresh all pivot tables in the workbook to ensure calculated values are up‑to‑date
                workbook.Worksheets.RefreshPivotTables();

                // Configure PDF save options to preserve the document structure (layout)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true
                };

                // Save the refreshed workbook as a PDF file
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine("Pivot tables refreshed and exported to PDF successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportRefreshedPivotTableToPdf.Run();
        }
    }
}
