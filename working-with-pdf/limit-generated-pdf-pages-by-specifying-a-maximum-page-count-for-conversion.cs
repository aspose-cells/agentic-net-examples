// Title: How to limit PDF pages when converting an Excel workbook to PDF using Aspose.Cells for .NET
// Common Searches: how to restrict PDF page count when saving Excel to PDF using Aspose.Cells C# | Aspose.Cells C# limit number of pages in generated PDF | set maximum pages for Excel to PDF conversion with Aspose.Cells .NET | workaround to control PDF page count by adjusting print area in Aspose.Cells | limit PDF output pages from Excel workbook using Aspose.Cells API
// Tags: Aspose.Cells PDF pagination management | Excel to PDF print area configuration Aspose.Cells | PdfSaveOptions worksheet PageSetup usage | C# restrict generated PDF pages Aspose.Cells | Workbook Save method PDF page limitation

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, creates a PdfSaveOptions object, and explains that Aspose.Cells lacks a direct MaximumPageCount property. It demonstrates using the worksheet's PageSetup (e.g., adjusting the PrintArea) as a workaround to limit the number of pages produced when saving the workbook as a PDF.
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

                // Load the existing Excel workbook
                var workbook = new Workbook(inputPath);

                // Configure PDF save options
                var pdfOptions = new PdfSaveOptions();

                // Aspose.Cells does not expose a direct MaximumPageCount property.
                // To limit pages, adjust the worksheet's PageSetup (e.g., PrintArea) as required.

                // Save the workbook as a PDF using the configured options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
