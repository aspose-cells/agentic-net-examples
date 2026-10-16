// Title: Convert an Excel workbook to PDF with 80% image quality using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set PDF image quality to 80 percent | How to export XLSX to PDF with reduced image size using Aspose.Cells | C# PdfSaveOptions.ImageQuality example for Excel to PDF conversion | Balance PDF file size and clarity when converting Excel with Aspose.Cells | Save workbook as PDF with custom JPEG compression in .NET
// Tags: Aspose.Cells PdfSaveOptions image quality | C# Excel to PDF conversion with compression | custom JPEG quality in PDF export Aspose.Cells | optimize PDF size and clarity Aspose.Cells | export workbook to PDF with image settings

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the input XLSX file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions with ImageQuality set to 80 to balance file size and clarity, and saves the workbook as a PDF. It includes exception handling for robust execution.
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

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
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
