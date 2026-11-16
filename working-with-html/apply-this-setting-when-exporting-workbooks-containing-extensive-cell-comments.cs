// Title: Export an Excel workbook to PDF while preserving cell comments using Aspose.Cells in C#
// Common Searches: how to keep Excel cell comments when converting to PDF with Aspose.Cells C# | Aspose.Cells C# export workbook to PDF include comments option | C# PdfSaveOptions comment visibility setting for PDF export | preserve Excel annotations in PDF using Aspose.Cells library | default behavior of Aspose.Cells PDF export for cell comments
// Tags: Aspose.Cells PDF export retain cell comments | C# PdfSaveOptions comment visibility | export Excel workbook to PDF with annotations | .NET convert Excel to PDF including comments | Aspose.Cells workbook to PDF comment handling | configure PDF save options for Excel comments

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook, creates a PdfSaveOptions object with default settings, and saves the workbook as a PDF file. It demonstrates how to export a workbook that contains extensive cell comments using Aspose.Cells for .NET while handling file‑not‑found and runtime exceptions.
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

                // Load the workbook containing cell comments
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF export options (default settings)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Export the workbook to PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
