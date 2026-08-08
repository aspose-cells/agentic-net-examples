// Title: Save Excel workbook as PDF with slicer visual style using Aspose.Cells for .NET
// Description: Loads an Excel file, sets PdfSaveOptions.ExportDocumentStructure to retain slicer formatting, and saves the workbook as a PDF while handling missing files and runtime errors.
// Keywords: Aspose.Cells | C# PDF export | ExportDocumentStructure | slicer visual style | Excel to PDF conversion | PdfSaveOptions example | .NET workbook to PDF | preserve slicer formatting
// Common Searches: Aspose.Cells keep slicer style when exporting to PDF | PdfSaveOptions ExportDocumentStructure C# example | convert Excel with slicers to PDF .NET | save workbook as PDF preserving slicer appearance | C# Aspose.Cells PDF options for slicers
// Developer Intent: Export an Excel workbook to PDF while embedding the slicer's visual appearance.
// Use Cases: Create printable PDFs of interactive dashboards that include slicers, ensuring the colors and layout match the original Excel view. | Automate batch conversion of financial models containing slicers to PDFs for archival without losing visual fidelity. | Generate client‑ready PDF reports from Excel templates with slicers, preserving exact styling for presentations.
// AI Prompts: Write C# code with Aspose.Cells that saves an Excel file as PDF and keeps slicer visual styles, showing the required PdfSaveOptions settings. | Explain how ExportDocumentStructure influences slicer rendering in PDF output and recommend any additional options for fine‑tuning appearance. | Provide a step‑by‑step checklist to verify that slicer formatting is correctly rendered in the generated PDF file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an Excel file, sets PdfSaveOptions.ExportDocumentStructure to retain slicer formatting, and saves the workbook as a PDF while handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains a slicer
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options to retain slicer visual styles
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true
                };

                // Save the workbook as PDF using the specified options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
