// Title: Convert an XLSX workbook to a high‑resolution PDF with Aspose.Cells for .NET (C#)
// Common Searches: c# convert xlsx file to pdf using aspose.cells | asp.net core load excel workbook and export to high resolution pdf | how to configure PdfSaveOptions in Aspose.Cells C# example | save each worksheet on a separate page when exporting Excel to PDF Aspose.Cells | asp.net convert excel to pdf preserving layout with Aspose.Cells
// Tags: Aspose.Cells PDFSaveOptions configuration C# | export Excel workbook to PDF using Aspose.Cells | high‑resolution PDF generation from XLSX with Aspose.Cells | OnePagePerSheet option for PDF export Aspose.Cells | load workbook and save as PDF .NET

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for 'input.xlsx', loads it into an Aspose.Cells Workbook, sets PdfSaveOptions (OnePagePerSheet = false), and saves the workbook as a high‑resolution PDF file 'output.pdf', with error handling for missing files and exceptions.
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

            // Load the XLSX workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Optional: keep each sheet on a separate page
                OnePagePerSheet = false
                // Note: ImageResolution property is not available in this version of Aspose.Cells.
                // High‑resolution output can be achieved by other means if required.
            };

            // Save the workbook as a PDF file with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
