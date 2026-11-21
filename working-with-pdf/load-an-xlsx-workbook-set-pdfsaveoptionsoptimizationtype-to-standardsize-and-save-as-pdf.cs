// Title: Convert an XLSX workbook to PDF with Standard size optimization using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# convert Excel file to PDF with Standard optimization | Set PdfSaveOptions OptimizationType to Standard in Aspose.Cells .NET | Save XLSX workbook as PDF using PdfOptimizationType.Standard in C# | How to apply standard size PDF optimization when exporting Excel to PDF with Aspose.Cells | Example code for converting input.xlsx to output.pdf with standard PDF size using Aspose.Cells
// Tags: Aspose.Cells PDF conversion Standard optimization | PdfSaveOptions OptimizationType Standard C# | Export Excel workbook to PDF Aspose.Cells | Standard size PDF output from XLSX Aspose.Cells | C# workbook.Save PDF options Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies that 'input.xlsx' exists, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions with OptimizationType set to PdfOptimizationType.Standard, and saves the workbook as 'output.pdf', handling any exceptions that may occur.
class Program
{
    static void Main()
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

            // Load the XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options and set optimization type
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use the standard optimization type (adjusted to the correct enum value)
                OptimizationType = PdfOptimizationType.Standard
            };

            // Save the workbook as a PDF file
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
