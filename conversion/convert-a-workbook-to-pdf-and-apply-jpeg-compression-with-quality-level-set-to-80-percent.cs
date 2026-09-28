// Title: How to export an Excel workbook to PDF with 80% JPEG compression using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.ImageCompression to JPEG, sets ImageQuality to 80, and saves the workbook as a PDF. | Create a console application that verifies the source Excel file, configures JPEG compression at 80% for PDF export, and writes the compressed PDF to a specified path. | Explain how to use Aspose.Cells PdfSaveOptions to reduce the size of a generated PDF by applying JPEG compression with a custom quality level in a .NET project.
// Common Searches: Aspose.Cells export Excel to PDF with JPEG compression 80 percent | C# set image quality when saving workbook as PDF using Aspose.Cells | How to reduce PDF file size from Excel conversion in .NET | PdfSaveOptions ImageCompression JPEG example Aspose.Cells | Adjust JPEG quality for PDF output from Excel workbook in C#
// Tags: Aspose.Cells PDF export JPEG compression | PdfSaveOptions image quality setting | C# Excel to PDF with image compression | JPEG quality control Aspose.Cells | optimize PDF size Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The program checks that the input Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, sets ImageCompression to JPEG and ImageQuality to 80, then saves the workbook as a compressed PDF while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
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
