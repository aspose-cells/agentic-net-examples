// Title: C# – Convert HTML to PDF (A4 Landscape) with Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, sets the worksheet to A4 landscape (optionally fit‑to‑page), and saves it as a PDF using PdfSaveOptions.
// Keywords: Aspose.Cells | HTML to PDF | C# | .NET | A4 landscape | PdfSaveOptions | page setup | worksheet export | convert HTML | PDF orientation
// Common Searches: Aspose.Cells convert html to pdf c# | set a4 landscape orientation aspose cells pdf | c# export html workbook to pdf with page setup | how to change page orientation before saving pdf aspose.cells | pdfsaveoptions a4 landscape c#
// Developer Intent: Create a PDF from an HTML file with A4 landscape orientation using Aspose.Cells in C#.
// Use Cases: Generate printable reports from web‑based HTML dashboards. | Batch convert HTML invoices to A4 landscape PDFs for archiving. | Produce documentation PDFs from HTML templates with a fixed landscape layout.
// AI Prompts: Write C# code that reads an HTML file, loads it into an Aspose.Cells Workbook, sets A4 landscape page setup, and saves it as a PDF. | Explain how to use PdfSaveOptions to fit an entire worksheet onto a single A4 landscape page in Aspose.Cells. | Show how to add error handling for missing input files and ensure the output directory exists when converting HTML to PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHtmlToPdf
{
    // Loads an HTML file into an Aspose.Cells Workbook, sets the worksheet to A4 landscape (optionally fit‑to‑page), and saves it as a PDF using PdfSaveOptions.
    class Program
    {
        static void Main()
        {
            const string inputPath = @"C:\Temp\input.html";
            const string outputPath = @"C:\Temp\output.pdf";

            try
            {
                // Verify that the input HTML file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the HTML file into a new workbook instance
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet created from the HTML
                Worksheet sheet = workbook.Worksheets[0];

                // Set page setup to A4 landscape
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                sheet.PageSetup.Orientation = PageOrientationType.Landscape;

                // Optional: fit the whole sheet on one page
                // sheet.PageSetup.FitToPagesWide = 1;
                // sheet.PageSetup.FitToPagesTall = 1;

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF file
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine("HTML has been converted to PDF with A4 landscape orientation.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
