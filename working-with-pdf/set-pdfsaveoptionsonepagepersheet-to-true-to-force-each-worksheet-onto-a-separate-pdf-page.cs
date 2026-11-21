// Title: How to export an Excel workbook to PDF with each worksheet on a separate page using Aspose.Cells for .NET
// Common Searches: how to set OnePagePerSheet in Aspose.Cells PDF export C# | Aspose.Cells C# export each Excel sheet to its own PDF page | convert .xlsx to PDF with one page per worksheet using Aspose.Cells | C# PdfSaveOptions OnePagePerSheet true example code | Aspose.Cells PDF conversion separate pages for each worksheet tutorial | save Excel workbook as PDF with sheet‑by‑sheet pagination in .NET
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | Excel to PDF conversion per worksheet C# | force one PDF page per sheet Aspose | save workbook as PDF with sheet pagination | Aspose.Cells PDF export sheet-level pagination

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, sets PdfSaveOptions.OnePagePerSheet to true, and saves the workbook as a PDF so that each worksheet appears on its own PDF page, with basic error handling for missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
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

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options to fit each sheet on a single page
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
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
