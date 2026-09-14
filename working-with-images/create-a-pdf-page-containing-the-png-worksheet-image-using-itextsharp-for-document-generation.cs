// Title: Generate a single-page PDF from the first worksheet of an Excel file using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures PdfSaveOptions to place each worksheet on a single PDF page, and saves the result as a PDF. | Extend the conversion program to output the PDF in landscape orientation and add a custom footer while preserving the one‑page‑per‑sheet layout. | Implement a reusable C# method that accepts input and output file paths, converts the first worksheet to PDF using Aspose.Cells, returns a boolean success flag, and logs any errors.
// Common Searches: Aspose.Cells C# export first worksheet to PDF single page per sheet | How to generate a PDF from Excel where each sheet fits on one page in .NET | C# PdfSaveOptions example to keep Excel sheet on one PDF page | Convert .xlsx to PDF using Aspose.Cells with page layout control
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet | C# export Excel worksheet to PDF | single-page PDF generation from Excel | Aspose.Cells workbook to PDF conversion | PDF output settings for Excel in .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input Excel file, loads it into an Aspose.Cells Workbook, applies PdfSaveOptions to force each worksheet onto a single PDF page, and saves the first worksheet as a PDF, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook (first worksheet) directly to PDF
            // OnePagePerSheet ensures each sheet fits on a single PDF page
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"PDF file successfully created at \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
