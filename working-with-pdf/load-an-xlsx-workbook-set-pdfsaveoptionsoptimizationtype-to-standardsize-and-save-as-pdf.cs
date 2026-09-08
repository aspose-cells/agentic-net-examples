// Title: Convert an XLSX workbook to PDF with Standard size optimization using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .xlsx file with Aspose.Cells, sets PdfSaveOptions.OptimizationType to Standard, and saves the result as a PDF file. | Create a .NET example that verifies the Excel file exists, applies standard PDF optimization via PdfSaveOptions, and outputs a PDF using Aspose.Cells.
// Common Searches: asp.net how to export Excel to PDF with standard optimization using Aspose.Cells | c# set PdfSaveOptions OptimizationType to Standard when converting XLSX to PDF | example code for saving workbook as PDF with standard size in Aspose.Cells | verify file exists before converting Excel to PDF with Aspose.Cells C#
// Tags: Aspose.Cells PDF optimization Standard | C# Excel to PDF conversion Aspose.Cells | PdfSaveOptions OptimizationType Standard example | Workbook.Save PDF Aspose.Cells | file existence check before Excel conversion C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Loads input.xlsx, checks its existence, configures PdfSaveOptions with OptimizationType = Standard, and saves the workbook as output.pdf using Aspose.Cells.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options with standard optimization
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OptimizationType = PdfOptimizationType.Standard
            };

            // Save the workbook as PDF using the specified options
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
