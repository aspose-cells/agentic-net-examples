// Title: Convert an Excel workbook to PDF using Aspose.Cells for .NET (default image DPI)
// Common Searches: how to convert an .xlsx file to PDF with Aspose.Cells in C# | Aspose.Cells PdfSaveOptions default image DPI when exporting Excel to PDF | C# code example for Excel to PDF conversion handling missing file errors | is it possible to set image resolution for PDF export in Aspose.Cells .NET | export workbook to PDF for high‑resolution printing using Aspose.Cells
// Tags: Aspose.Cells PDF conversion C# | PdfSaveOptions image DPI limitation | Excel to PDF high‑resolution export | Workbook.Save PDF Aspose.Cells | file existence validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file (input.xlsx) into an Aspose.Cells Workbook, creates a PdfSaveOptions object (noting that custom image DPI is not supported in the current version), and saves the workbook as a PDF (output.pdf) while checking for file existence and handling exceptions.
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

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (high‑resolution image DPI is not supported in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the specified options
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
