// Title: Convert an XLS workbook to PDF with Standard size optimization using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xls file, configures PdfSaveOptions for Standard optimization, and saves the workbook as a PDF with Aspose.Cells. | Show how to apply standard PDF optimization when exporting an Excel workbook to PDF in a .NET application using Aspose.Cells.
// Common Searches: Aspose.Cells C# set PdfSaveOptions for standard PDF optimization during export | How to export XLS to PDF with standard size optimization using Aspose.Cells .NET | PdfSaveOptions example for standard PDF output from Excel workbook | Saving an Excel workbook as PDF with standard optimization in a .NET project | C# code sample for PDF optimization type Standard in Aspose.Cells
// Tags: Aspose.Cells PDF standard optimization | C# PdfSaveOptions OptimizationType Standard | Export XLS to PDF Aspose.Cells | Standard size PDF output Excel .NET | Workbook.Save PDF options Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an existing XLS workbook, configures PdfSaveOptions with OptimizationType set to Standard, and saves the workbook as a PDF file, including input validation and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the optimization type to Standard (available enum value)
                OptimizationType = PdfOptimizationType.Standard
            };

            // Save the workbook as a PDF using the specified options
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
