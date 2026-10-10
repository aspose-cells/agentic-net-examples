// Title: Limit PDF pages per worksheet when converting an Excel workbook to PDF using Aspose.Cells for .NET (OnePagePerSheet option)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.OnePagePerSheet to true to ensure each worksheet is rendered on a single PDF page, and saves the workbook as a PDF. | Show how to configure Aspose.Cells PdfSaveOptions in a .NET application to restrict the number of PDF pages generated per worksheet during Excel‑to‑PDF conversion.
// Common Searches: asp.net aspose.cells limit number of PDF pages per sheet during Excel to PDF conversion | c# set OnePagePerSheet option in PdfSaveOptions to reduce PDF page count | how to generate single-page PDF for each worksheet using Aspose.Cells | restrict PDF output pages when saving workbook as PDF with Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# limit PDF pages per worksheet | Excel to PDF conversion page count restriction | Aspose.Cells PDF page limit .NET

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the existence of input.xlsx, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions with OnePagePerSheet enabled to limit each worksheet to a single PDF page, saves the result as output.pdf, and reports success or any errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (e.g., fit each sheet on one page)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

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
