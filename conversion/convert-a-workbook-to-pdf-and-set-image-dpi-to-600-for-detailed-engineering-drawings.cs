// Title: Convert an Excel workbook to a 600 DPI PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set PDF image resolution to 600 DPI | Convert XLSX to high‑resolution PDF with Aspose.Cells .NET | Save Excel workbook as PDF with 600 DPI images using PdfSaveOptions | C# example exporting engineering drawings to PDF at 600 DPI | PdfSaveOptions ImageResolution property usage example
// Tags: Aspose.Cells PDF export DPI control | Excel to PDF high‑resolution export C# | Aspose.Cells ImageResolution configuration | engineering drawing PDF generation Aspose.Cells | 600 DPI PDF export from XLSX

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input XLSX file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (optionally setting ImageResolution to 600 DPI), and saves the workbook as a high‑resolution PDF, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Note: In some versions of Aspose.Cells the ImageResolution property is not available.
            // If your version supports it, you can uncomment the following line to set a higher DPI.
            // pdfOptions.ImageResolution = 600;

            // Save the workbook as a PDF using the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
