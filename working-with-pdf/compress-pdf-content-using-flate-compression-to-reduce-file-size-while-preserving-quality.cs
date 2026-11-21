// Title: Save an Excel workbook as a Flate‑compressed PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# save workbook as PDF with Flate compression | how to reduce PDF size when exporting Excel with Aspose.Cells .NET | set PdfSaveOptions compression to Flate in C# example | C# code to export .xlsx to a compressed PDF using Aspose.Cells | Flate compression for PDF output from an Aspose.Cells workbook
// Tags: Aspose.Cells PDF compression Flate | C# Excel to PDF conversion Aspose.Cells | PdfSaveOptions compression property | reduce PDF file size Aspose.Cells | export workbook as compressed PDF .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the source .xlsx file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions instance (using the default Flate compression), and saves the workbook as a PDF, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source Excel file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default compression)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
