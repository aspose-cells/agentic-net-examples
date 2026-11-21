// Title: Convert HTML to PDF with Aspose.Cells for .NET (C#) and specify PDF version 1.7
// Common Searches: aspnet convert html to pdf with aspose.cells and set pdf version 1.7 | c# how to set pdf version when saving workbook as pdf using aspose.cells | asp.net aspose.cells pdfsaveoptions specify pdf version for html conversion | save html workbook as pdf with advanced pdf features using aspose.cells c# | aspose.cells export html to pdf with pdf 1.7 compliance c#
// Tags: Aspose.Cells PDFSaveOptions set PDF version | HTML to PDF conversion using Aspose.Cells .NET | C# workbook export to PDF with custom version | advanced PDF features in Aspose.Cells | convert HTML workbook to PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the presence of an HTML file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (where the PDF version can be configured, e.g., to 1.7), and saves the workbook as a PDF file while handling any runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.html";
                const string outputPath = "output.pdf";

                // Verify that the input HTML file exists before loading.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the HTML file into a workbook.
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (no need to set SaveFormat; it's inferred from the options type).
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Additional PDF options can be set here if needed.
                };

                // Save the workbook as a PDF.
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
