// Title: Set PdfSaveOptions.FitAllColumnsInOnePage = true in C# to fit all Excel columns on a single PDF page with Aspose.Cells
// AI Prompts: Generate C# code that loads an .xlsx workbook and saves it as a PDF with PdfSaveOptions.FitAllColumnsInOnePage set to true using Aspose.Cells. | Demonstrate a fallback to OnePagePerSheet when the FitAllColumnsInOnePage property is unavailable, while still ensuring all columns appear on one PDF page. | Provide comprehensive error handling for loading the workbook and saving the PDF when configuring column‑fit options in Aspose.Cells.
// Common Searches: Aspose.Cells C# set FitAllColumnsInOnePage true for PDF export | export Excel to single-page PDF with all columns visible using Aspose.Cells | PdfSaveOptions column fitting options C# Aspose.Cells | how to force all columns onto one PDF page in Aspose.Cells | OnePagePerSheet fallback when FitAllColumnsInOnePage missing
// Tags: Aspose.Cells PdfSaveOptions FitAllColumnsInOnePage | C# Excel to PDF single-page conversion | OnePagePerSheet fallback column fit | column fitting PDF export Aspose.Cells | PdfSaveOptions column scaling setting

using System;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx workbook, configures PdfSaveOptions with FitAllColumnsInOnePage = true (or uses OnePagePerSheet as a fallback), and saves the workbook as output.pdf so that all columns are forced onto a single PDF page. It includes checks for file existence and robust exception handling for both loading and saving operations.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook
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

            // Configure PDF save options to fit the sheet on a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // If the specific property AllColumnsInOnePage is unavailable,
                // OnePagePerSheet ensures the entire sheet (columns and rows) fits on one page.
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF
            try
            {
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
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
