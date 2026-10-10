// Title: Generate a PDF with a separate page for each worksheet using Aspose.Cells PdfSaveOptions.OnePagePerSheet in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.OnePagePerSheet to true, and saves each worksheet as an individual PDF page. | Show how to add file‑existence validation and exception handling when converting an Excel workbook to PDF with one page per sheet using Aspose.Cells. | Demonstrate configuring PdfSaveOptions for PDF export of multiple worksheets in a single PDF document in C#.
// Common Searches: asp.net convert excel workbook to pdf with one page per worksheet using aspose.cells | c# pdfsaveoptions onepagepersheet example | how to export each worksheet to separate pdf page with aspose cells | check if excel file exists before saving as pdf in c# asp.net | asp.net core aspose cells pdf export multiple sheets separate pages
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# export Excel to PDF separate worksheet pages | file existence validation before Aspose.Cells conversion | exception handling Aspose.Cells PDF export | multi-sheet PDF generation with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Loads an Excel workbook, verifies the source file, configures PdfSaveOptions.OnePagePerSheet = true, and saves the workbook as a PDF where each worksheet appears on its own page, with basic error handling.
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

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Set PDF save options: one page per worksheet
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Save the workbook as a PDF using the configured options
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
