// Title: How to export an Excel workbook to PDF with each worksheet on a separate page using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file and saves it as a PDF with one page per worksheet using Aspose.Cells. | Show how to configure PdfSaveOptions for OnePagePerSheet and apply it when calling Workbook.Save in a .NET application.
// Common Searches: how to save Excel as PDF with one page per sheet using Aspose.Cells .NET | Aspose.Cells PdfSaveOptions OnePagePerSheet example C# | convert .xlsx to .pdf with Aspose.Cells SaveFormat.Pdf | set PDF/A compliance when exporting workbook to PDF using Aspose.Cells
// Tags: per-sheet pagination option | Aspose.Cells Excel to PDF conversion | C# workbook PDF export | PDF/A level option | separate PDF page per worksheet

using System;
using System.IO;
using Aspose.Cells;

// The program checks for an input.xlsx file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions to place each worksheet on its own PDF page (OnePagePerSheet = true), and saves the result as output.pdf while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF export options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                // Export each worksheet to a separate page
                OnePagePerSheet = true
                // PDF/A compliance can be set here if supported by the library version
                // Compliance = PdfCompliance.PdfA1b
            };

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfSaveOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
