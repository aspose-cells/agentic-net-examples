// Title: How to include row and column headings on every PDF page when converting an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, enables printing of row and column headings for each worksheet, and saves the workbook as a PDF using Aspose.Cells PdfSaveOptions. | Show a try‑catch example that checks the existence of the input Excel file, activates worksheet headings, and exports the workbook to PDF with default PdfSaveOptions in a .NET console app. | Demonstrate how to apply page‑setup heading settings to all worksheets before calling Workbook.Save with PdfSaveOptions in a C# application.
// Common Searches: asp.net convert excel to pdf with row and column headings on each page using aspose.cells | c# enable print headings for pdf export in aspose.cells workbook | how to set worksheet page setup printheadings true before saving as pdf in .net | asp.net core aspose.cells pdfsaveoptions print headings each page | excel to pdf conversion include row numbers and column letters as headings asp.net
// Tags: Aspose.Cells PDF export with worksheet headings | C# enable row and column headings in PDF conversion | Excel workbook page setup print headings Aspose | PdfSaveOptions default usage Aspose.Cells | Convert .xlsx to PDF including headings .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, turns on printing of row and column headings for every worksheet, and saves the file as a PDF using Aspose.Cells with default PdfSaveOptions, ensuring each PDF page shows the worksheet headings.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Enable printing of row and column headings for each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PrintHeadings = true;
            }

            // Configure PDF save options (default options are sufficient here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file with the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
