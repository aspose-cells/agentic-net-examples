// Title: Set PdfSaveOptions.OptimizationType to MinimumSize and convert an XLSX workbook to PDF using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an XLSX file with Aspose.Cells, configures PdfSaveOptions.OptimizationType to MinimumSize, and saves the workbook as a PDF. | Show how to minimize the generated PDF size when exporting an Excel workbook to PDF by applying the MinimumSize optimization setting in Aspose.Cells.
// Common Searches: Aspose.Cells C# set PDF optimization type to MinimumSize | How to export Excel to PDF with the smallest file size using Aspose.Cells | PdfSaveOptions MinimumSize example for XLSX to PDF conversion | Reduce PDF size when converting a workbook to PDF in .NET with Aspose.Cells | C# code sample for PdfOptimizationType.MinimumSize in Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions MinimumSize | C# Excel to PDF conversion optimization | PdfOptimizationType MinimumSize usage | reduce PDF size Aspose.Cells | export XLSX as PDF with size reduction

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that the input XLSX file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, sets its OptimizationType to MinimumSize to produce the smallest possible PDF, saves the workbook as a PDF file, and handles any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default optimization)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
