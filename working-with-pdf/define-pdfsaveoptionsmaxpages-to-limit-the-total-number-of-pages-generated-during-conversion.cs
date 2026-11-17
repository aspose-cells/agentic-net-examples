// Title: Limit the number of pages exported from an Excel workbook to PDF using PdfSaveOptions in Aspose.Cells for .NET
// Common Searches: how to set maximum pages when converting Excel to PDF with Aspose.Cells C# | Aspose.Cells PdfSaveOptions limit PDF output to first 10 pages | C# restrict PDF page count during workbook save using Aspose.Cells | configure page count limit for Excel to PDF conversion in .NET
// Tags: PdfSaveOptions page count limit | Aspose.Cells limit PDF pages | Excel to PDF conversion page restriction .NET | C# workbook save as PDF with max pages | Aspose.Cells PDF export page control

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a PdfSaveOptions object with the PageCount property set to 10 to cap the PDF output to ten pages, ensures the target directory exists, and saves the workbook as a PDF. It also includes checks for the source file and basic exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the source workbook exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (limit to first 10 pages)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // StartPage property is not available; default is the first page
                PageCount = 10
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
