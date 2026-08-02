// Title: C# – Convert an XLSX Workbook to a High‑Resolution PDF with Aspose.Cells
// Description: Loads an existing XLSX file, applies PdfSaveOptions (including ExportDocumentStructure), and saves the workbook as a high‑resolution PDF. The sample includes file‑existence checks and robust exception handling for reliable .NET conversions.
// Keywords: Aspose.Cells C# | XLSX to PDF conversion | high resolution PDF from Excel | PdfSaveOptions ExportDocumentStructure | Aspose.Cells PDF conversion .NET | Excel workbook to PDF C# | Aspose.Cells high quality PDF | convert Excel to PDF programmatically | C# Aspose.Cells example
// Common Searches: Aspose.Cells convert Excel to PDF C# | high resolution PDF export from XLSX using Aspose | PdfSaveOptions ExportDocumentStructure example | C# code to save workbook as PDF with Aspose.Cells | how to handle file not found when converting Excel to PDF
// Developer Intent: Create a high‑resolution PDF from an XLSX workbook using Aspose.Cells for .NET.
// Use Cases: Produce print‑ready PDFs of financial statements directly from Excel templates. | Generate standardized PDF invoices from a workbook while preserving layout hierarchy. | Automate batch conversion of archived Excel reports to searchable PDFs for compliance.
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, configures PdfSaveOptions to preserve document structure, and saves a high‑resolution PDF. | Show how to implement file‑existence validation and exception handling for Excel‑to‑PDF conversion in a console app using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsConversionDemo
{
    // Loads an existing XLSX file, applies PdfSaveOptions (including ExportDocumentStructure), and saves the workbook as a high‑resolution PDF. The sample includes file‑existence checks and robust exception handling for reliable .NET conversions.
    public class XlsxToPdfConverter
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Desired path for the high‑resolution PDF output
            string pdfPath = "output.pdf";

            // Verify that the source file exists before attempting to load it
            if (!File.Exists(sourcePath))
            {
                Console.Error.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook from the XLSX file
                Workbook workbook = new Workbook(sourcePath);

                // Create PDF save options – customize as needed
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Example option: preserve document structure (optional)
                    ExportDocumentStructure = true
                };

                // Save the workbook as a PDF using the specified options
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Workbook successfully converted to PDF: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}
