// Title: Generate a PDF with a separate page for each worksheet using Aspose.Cells PdfSaveOptions.OnePagePerSheet in C#
// AI Prompts: Write C# code that loads an .xlsx workbook, sets PdfSaveOptions.OnePagePerSheet to true, and saves it as a PDF where each worksheet appears on its own page. | Show how to configure Aspose.Cells PdfSaveOptions to produce a multi‑page PDF, one page per sheet, and persist the workbook in a .NET application.
// Common Searches: how to export each Excel sheet to a separate PDF page using Aspose.Cells C# | Aspose.Cells PdfSaveOptions OnePagePerSheet property example | C# convert workbook to PDF with one page per worksheet Aspose | save Excel file as multi‑page PDF per sheet Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# export Excel to PDF per worksheet | generate multi-page PDF from workbook Aspose | set per-sheet pagination PDF Aspose.Cells | save workbook as PDF with sheet separation .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads an existing Excel workbook, enables PdfSaveOptions.OnePagePerSheet, and saves the file as a PDF where each worksheet is rendered on a separate page.
    class Program
    {
        static void Main(string[] args)
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

                // Configure PDF save options: one page per worksheet
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Save the workbook as PDF using the configured options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
