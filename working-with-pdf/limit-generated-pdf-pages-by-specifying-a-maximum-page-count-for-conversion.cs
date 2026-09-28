// Title: How to limit exported pages when converting an Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to save an Excel workbook as a PDF while restricting the output to a defined maximum number of pages. | Illustrate configuring Aspose.Cells PdfSaveOptions to export only the first N pages of a workbook during PDF conversion.
// Common Searches: Aspose.Cells C# export first 5 pages of Excel to PDF | Set maximum page count when saving workbook as PDF using Aspose.Cells | Configure PdfSaveOptions to limit PDF pages in .NET conversion | C# convert Excel to PDF with page count restriction Aspose
// Tags: Aspose.Cells PDF page limit setting | C# PDF page count configuration with Aspose | Excel to PDF conversion limit pages Aspose | Aspose.Cells set PageIndex zero based | restrict PDF output pages Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the source Excel file, loads it into a Workbook, sets PdfSaveOptions with PageIndex = 0 and PageCount = 5 to export only the first five pages, saves the workbook as a PDF, and handles any exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the existing Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options to export only the first 5 pages
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // PageIndex is zero‑based; PageCount specifies how many pages to export
                    PageIndex = 0,
                    PageCount = 5
                };

                // Save the workbook as a PDF using the configured options
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
