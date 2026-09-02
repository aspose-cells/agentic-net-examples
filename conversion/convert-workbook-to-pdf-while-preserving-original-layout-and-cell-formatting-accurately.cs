// Title: Convert an Excel workbook to PDF while preserving layout and cell formatting using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, calculates all formulas, and saves it as a PDF using PdfSaveOptions configured to retain the original worksheet layout. | Show how to set PdfSaveOptions properties such as ExportDocumentStructure, CheckWorkbookDefaultFont, OnePagePerSheet, and AllColumnsInOnePagePerSheet to keep formatting during Excel‑to‑PDF conversion. | Create a robust conversion routine that verifies the source file exists, handles exceptions, and logs success or error messages.
// Common Searches: Aspose.Cells preserve cell formatting when exporting Excel to PDF in C# | How to keep original worksheet layout during Excel to PDF conversion with Aspose.Cells .NET | PdfSaveOptions settings for maintaining fonts and bookmarks in Excel to PDF export | Convert .xlsx to PDF without forcing one page per sheet using Aspose.Cells | C# code example for Excel to PDF conversion with formula calculation using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions document structure | Formula calculation before PDF export in Aspose.Cells | Default font handling in Excel to PDF conversion | C# pre‑conversion file existence validation | Prevent single-page-per-sheet in Aspose.Cells PDF output

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Converts an Excel file (input.xlsx) to a PDF (output.pdf) using Aspose.Cells in C#. The code verifies the source file, loads the workbook, calculates formulas, configures PdfSaveOptions to preserve document structure, use the workbook's default font, and avoid forcing a single page per sheet, then saves the PDF.
    public class WorkbookToPdfConverter
    {
        public static void Run()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Path where the PDF will be saved
            string destPath = "output.pdf";

            try
            {
                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook from the file system
                Workbook workbook = new Workbook(sourcePath);

                // Ensure that all formulas are calculated before conversion
                workbook.CalculateFormula();

                // Configure PDF save options to preserve layout and formatting
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Keep the document structure (e.g., headings, bookmarks)
                    ExportDocumentStructure = true,

                    // Use the workbook's default font when a specific font is missing
                    CheckWorkbookDefaultFont = true,

                    // Do not force all content onto a single page per sheet
                    OnePagePerSheet = false,

                    // Keep each column on its own page only if explicitly required
                    AllColumnsInOnePagePerSheet = false
                };

                // Save the workbook as a PDF using the configured options
                workbook.Save(destPath, pdfOptions);

                Console.WriteLine("Workbook successfully converted to PDF while preserving layout and formatting.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookToPdfConverter.Run();
        }
    }
}
