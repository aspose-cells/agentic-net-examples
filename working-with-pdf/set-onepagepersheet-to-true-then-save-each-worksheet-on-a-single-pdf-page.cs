// Title: Export an Excel workbook to PDF with each worksheet forced onto a single page using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.OnePagePerSheet to true, and saves the workbook as a PDF. | Show how to configure Aspose.Cells PdfSaveOptions so that every worksheet is rendered on one PDF page before calling Workbook.Save.
// Common Searches: Aspose.Cells C# how to export each worksheet to a single PDF page | PdfSaveOptions OnePagePerSheet property usage example | Save Excel workbook as PDF with one page per sheet in .NET | Force all worksheets onto one page per sheet when converting Excel to PDF with Aspose | C# Aspose.Cells PDF export page layout settings
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# export Excel to single-page PDF | Aspose.Cells worksheet PDF page layout | PDF conversion one page per sheet Aspose | Aspose.Cells save workbook as PDF with page control

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample checks for the input Excel file, loads it into an Aspose.Cells Workbook, enables PdfSaveOptions.OnePagePerSheet to force each worksheet onto a single PDF page, and saves the workbook as a PDF while handling any exceptions.
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to force each worksheet onto a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
