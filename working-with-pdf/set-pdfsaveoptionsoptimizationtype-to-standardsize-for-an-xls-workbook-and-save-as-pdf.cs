// Title: How to set PdfSaveOptions.OptimizationType to StandardSize and convert an XLS workbook to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set PdfSaveOptions OptimizationType to StandardSize before PDF export | Convert XLS file to PDF with StandardSize optimization in Aspose.Cells | Example of configuring PdfSaveOptions for StandardSize PDF output in .NET | C# code to export Excel workbook to PDF using StandardSize optimization | Aspose.Cells PDF export with size optimization for large spreadsheets
// Tags: PdfSaveOptions StandardSize setting | XLS to PDF conversion Aspose.Cells | Aspose.Cells PDF export options .NET | Excel workbook PDF size optimization C# | C# Aspose.Cells PDF optimization

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an XLS workbook with Aspose.Cells, creates a PdfSaveOptions object, sets its OptimizationType to StandardSize, and saves the workbook as a PDF, demonstrating how to apply size‑optimized PDF export in a .NET application.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xls";
                string outputPath = "output.pdf";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing XLS workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default optimization)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

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
