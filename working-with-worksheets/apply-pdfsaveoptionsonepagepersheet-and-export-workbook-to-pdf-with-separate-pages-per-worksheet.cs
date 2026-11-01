// Title: Save an Excel workbook as a PDF with one page per worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export workbook to PDF with each sheet on a separate page | PdfSaveOptions OnePagePerSheet property usage example in .NET | How to generate PDF where each Excel worksheet becomes its own page using Aspose.Cells | C# code to convert .xlsx to PDF with per‑sheet pagination via Aspose | Set one page per sheet when saving Excel to PDF with Aspose.Cells library
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# Excel to PDF per worksheet | PDF export with sheet-wise pagination Aspose | Aspose.Cells workbook.Save PDF options | one page per sheet PDF conversion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing .xlsx file into an Aspose.Cells Workbook, configures PdfSaveOptions with OnePagePerSheet = true, and saves the workbook as a PDF so that each worksheet appears on its own page, including basic file‑existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Set PDF save options: one page per worksheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
