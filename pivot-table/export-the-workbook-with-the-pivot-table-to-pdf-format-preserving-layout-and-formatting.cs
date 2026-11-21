// Title: Export an Excel workbook containing a pivot table to PDF while preserving layout using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export workbook with pivot table to PDF preserving column widths | How to keep pivot table formatting when saving Excel as PDF with Aspose.Cells | Aspose.Cells PdfSaveOptions settings to avoid scaling and retain original layout in .NET | Saving an Excel file that includes a pivot table to PDF without forcing one page per sheet | Export Excel pivot table to PDF using Aspose.Cells while maintaining original layout
// Tags: Aspose.Cells PDFSaveOptions layout preservation | C# export pivot table to PDF | Aspose.Cells preserve column widths PDF | Excel to PDF conversion with pivot table Aspose | disable one-page-per-sheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing Excel workbook that contains a pivot table, configures PdfSaveOptions to prevent one-page-per-sheet scaling and to keep column widths, ensures the output directory exists, and saves the workbook as a PDF while retaining the original layout and formatting.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the pivot table
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // NOTE: PrintGridlines property is not available in the current Aspose.Cells version.
            // If needed, enable gridlines via other means (e.g., page setup) in future revisions.

            // Configure PDF save options
            var pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false,               // Do not force each sheet onto a single page
                AllColumnsInOnePagePerSheet = false    // Preserve column widths without scaling
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the workbook (including the pivot table) to PDF
            try
            {
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save PDF: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
