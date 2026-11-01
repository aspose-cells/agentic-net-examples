// Title: How to convert an Excel .xlsx workbook to PDF/A‑1b in C# while keeping original layout and cell formatting with Aspose.Cells
// Common Searches: C# Aspose.Cells convert XLSX to PDF/A-1b preserving column widths | save Excel workbook as PDF without scaling columns using Aspose.Cells .NET | Aspose.Cells PDFSaveOptions keep original sheet layout during conversion | export Excel to multi‑page PDF while retaining cell formatting in C# | Aspose.Cells prevent all columns on one page when saving to PDF
// Tags: Aspose.Cells PDFSaveOptions layout control | C# Excel to PDF/A-1b export | keep original column widths Aspose.Cells | multiple pages per sheet PDF conversion | cell formatting retention PDF export .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the input .xlsx file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to disable column scaling (AllColumnsInOnePagePerSheet = false), allow multiple pages per sheet (OnePagePerSheet = false), and set PDF/A‑1b compliance, then saves the workbook as a PDF while preserving the original layout and cell formatting.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to preserve layout and formatting
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not force all columns onto a single page; keep original column widths
                AllColumnsInOnePagePerSheet = false,
                // Allow multiple pages per sheet if the content exceeds one page
                OnePagePerSheet = false,
                // Use PDF/A-1b compliance to preserve formatting details
                Compliance = PdfCompliance.PdfA1b
                // Note: ImageResolution property is not available in this version of Aspose.Cells
            };

            // Save the workbook as a PDF file with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF file saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
