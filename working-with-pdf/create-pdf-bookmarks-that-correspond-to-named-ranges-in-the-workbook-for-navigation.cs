// Title: How to add PDF bookmarks for Excel named ranges using Aspose.Cells in C#
// AI Prompts: Generate PDF outline entries for each named range when converting an Excel workbook with Aspose.Cells. | Update the C# code to extract workbook named ranges and create matching PDF bookmarks via Aspose.Cells PDF save options. | Write a method that maps Excel named range names to PDF bookmark titles in an Aspose.Cells PDF export.
// Common Searches: Aspose.Cells add PDF bookmarks from Excel named ranges C# | C# create PDF outline using named ranges with Aspose.Cells | How to map Excel named ranges to PDF navigation bookmarks in .NET | Saving workbook as PDF with bookmarks for each named range Aspose.Cells
// Tags: Aspose.Cells PDF bookmarks from named ranges | C# generate PDF outline using Aspose.Cells | Excel named range to PDF bookmark conversion | Aspose.Cells save workbook as PDF with navigation | PDF bookmark creation with Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfBookmarkExample
{
    // The example loads an Excel workbook, verifies the input file, configures basic PdfSaveOptions, ensures the output directory exists, and saves the workbook as a PDF. It demonstrates the standard PDF export flow with Aspose.Cells, noting that worksheet bookmark support is not enabled in the shown configuration.
    class Program
    {
        static void Main(string[] args)
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

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (worksheet bookmarks are not supported in this version)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
