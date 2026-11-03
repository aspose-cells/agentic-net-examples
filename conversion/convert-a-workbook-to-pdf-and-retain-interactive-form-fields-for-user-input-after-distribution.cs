// Title: Convert an Excel workbook to a fillable PDF while preserving form fields using Aspose.Cells for .NET
// Common Searches: how to generate a fillable PDF from an .xlsx file with Aspose.Cells C# | preserve Excel form controls when saving as PDF using Aspose.Cells .NET | Aspose.Cells PdfSaveOptions keep interactive fields in exported PDF | C# code to export workbook to PDF with fillable form fields | convert Excel to PDF without losing data validation drop‑downs using Aspose
// Tags: Aspose.Cells PDFSaveOptions fillable PDF | export Excel workbook to PDF with interactive fields C# | keep Excel form controls in PDF export using Aspose | convert .xlsx to fillable PDF .NET | retain data validation drop-downs in PDF conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // The example verifies the source Excel file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (default settings preserve fillable form fields), and saves the workbook as a PDF, handling any runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the existing Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default options are sufficient for preserving form fields)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF with the specified options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
