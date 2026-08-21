// Title: Convert an Excel workbook to PDF while preserving layout, formatting, and pagination with Aspose.Cells for .NET
// Description: This C# example demonstrates how to load an .xlsx file using Aspose.Cells, configure PdfSaveOptions (ExportDocumentStructure, CheckWorkbookDefaultFont, OnePagePerSheet) to keep the original workbook appearance, and save it as a PDF. The code also includes file‑existence checks and exception handling for robust conversion.
// Keywords: Aspose.Cells | Excel to PDF conversion | preserve Excel layout | preserve cell formatting | PDF pagination | OnePagePerSheet | ExportDocumentStructure | CheckWorkbookDefaultFont | .NET PDF export | C# Aspose.Cells example | multi‑sheet PDF | Excel bookmarks PDF
// Common Searches: Aspose.Cells keep Excel formatting when saving as PDF | PDF conversion options to retain page breaks in Aspose.Cells | How to export each worksheet to a separate PDF page using Aspose.Cells .NET | Enable bookmarks in PDF generated from Excel with Aspose.Cells | C# code for Excel to PDF with layout preservation
// Developer Intent: Generate a PDF from an Excel workbook that looks identical to the source, including fonts, cell styles, page breaks, and document structure.
// Use Cases: Produce printable PDF reports from financial workbooks where each sheet starts on a new page. | Create searchable PDFs with bookmarks for multi‑sheet project documentation. | Integrate Excel‑to‑PDF conversion into a web service while ensuring fonts and cell formatting are retained.
// AI Prompts: Write C# code using Aspose.Cells to convert an .xlsx file to PDF with ExportDocumentStructure, CheckWorkbookDefaultFont, and OnePagePerSheet enabled. | Explain how ExportDocumentStructure and CheckWorkbookDefaultFont affect the visual fidelity of PDFs generated from Excel files. | Show best practices for handling missing source files and logging errors during Excel‑to‑PDF conversion with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to load an .xlsx file using Aspose.Cells, configure PdfSaveOptions (ExportDocumentStructure, CheckWorkbookDefaultFont, OnePagePerSheet) to keep the original workbook appearance, and save it as a PDF. The code also includes file‑existence checks and exception handling for robust conversion.
    public class WorkbookToPdfConverter
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source Excel workbook
            string sourceFile = "input.xlsx";

            // Desired path for the output PDF file
            string pdfFile = "output.pdf";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Error: Source file '{sourceFile}' not found.");
                return;
            }

            try
            {
                // Load the workbook from the file
                Workbook workbook = new Workbook(sourceFile);

                // Configure PDF save options to retain layout and formatting
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Preserve document structure such as bookmarks and headings
                    ExportDocumentStructure = true,

                    // Use the workbook's default font when a specific font is missing
                    CheckWorkbookDefaultFont = true,

                    // Ensure each worksheet starts on a new page (helps keep original pagination)
                    OnePagePerSheet = true
                };

                // Save the workbook as PDF with the specified options
                workbook.Save(pdfFile, pdfOptions);

                Console.WriteLine($"Conversion completed: '{sourceFile}' → '{pdfFile}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
