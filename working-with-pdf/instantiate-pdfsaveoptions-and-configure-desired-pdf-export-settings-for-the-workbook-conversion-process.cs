// Title: How to use Aspose.Cells PdfSaveOptions to export an Excel workbook as a single-page PDF in C#
// AI Prompts: Write C# code that loads an .xlsx file, creates a PdfSaveOptions instance, enables the single-page-per-sheet mode, and saves the workbook as a PDF using Aspose.Cells. | Show how to modify additional PDF export parameters such as PDF/A compliance level and image DPI through PdfSaveOptions before invoking Workbook.Save in a .NET application.
// Common Searches: Aspose.Cells C# enable single page per sheet PDF export | Convert Excel workbook to PDF with each sheet on one page using Aspose.Cells | PdfSaveOptions property to force one page per worksheet in .NET | Set image DPI for PDF output in Aspose.Cells C# example | C# sample for saving workbook as PDF with custom options in Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# Excel to PDF conversion Aspose.Cells | adjust PDF raster resolution Aspose.Cells | Workbook.Save with custom PdfSaveOptions C# | Aspose.Cells PDF output configuration example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // This program checks for the input XLSX file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to produce a single-page-per-sheet PDF, and saves the result as a PDF file.
    class Program
    {
        static void Main()
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

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                    // Additional options (e.g., compliance, image resolution) can be set here
                    // if supported by the version of Aspose.Cells being used.
                };

                // Save the workbook as a PDF file
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
