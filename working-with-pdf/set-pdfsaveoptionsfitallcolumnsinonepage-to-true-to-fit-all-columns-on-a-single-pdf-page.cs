// Title: How to fit all Excel columns onto a single PDF page using Aspose.Cells PdfSaveOptions in C#
// AI Prompts: Write C# code that loads an .xlsx workbook, sets PdfSaveOptions.FitAllColumnsInOnePage (or OnePagePerSheet) to true, and saves the workbook as a PDF. | Show how to configure Aspose.Cells PdfSaveOptions so that every column of a worksheet is forced onto one PDF page while keeping the sheet data intact.
// Common Searches: Aspose.Cells C# fit all columns on one PDF page example | PdfSaveOptions FitAllColumnsInOnePage property usage | Export Excel sheet to single-page PDF with Aspose.Cells | C# convert workbook to PDF with columns forced onto one page
// Tags: Aspose.Cells PdfSaveOptions column fitting | C# export Excel to single-page PDF | PdfSaveOptions OnePagePerSheet property | fit all columns PDF Aspose.Cells | single-page PDF export from workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an existing Excel workbook, creates a PdfSaveOptions object with OnePagePerSheet (which fits all columns on a single PDF page), ensures the output folder exists, and saves the workbook as a PDF file.
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

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to fit the entire sheet on a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true // Fits all columns (and rows) on one page
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
