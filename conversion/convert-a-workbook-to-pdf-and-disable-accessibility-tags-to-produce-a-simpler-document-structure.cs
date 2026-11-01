// Title: Convert an Excel workbook to PDF without creating tagged PDFs using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF without accessibility tags | Generate a non‑tagged PDF from an .xlsx file using Aspose.Cells | C# save workbook as PDF without PDF/UA compliance in Aspose.Cells | Disable PDF accessibility tags when converting Excel to PDF with Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions non‑tagged PDF | C# Excel workbook PDF conversion Aspose | remove PDF/UA tags Aspose.Cells output | simple PDF generation from Excel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, uses the default PdfSaveOptions (which do not generate tagged PDFs), ensures the output folder exists, and saves the workbook as a PDF file, handling any errors that may occur.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.pdf";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputFile);

            // Configure PDF save options (default options do not create tagged PDFs)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF using the configured options
            workbook.Save(outputFile, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
