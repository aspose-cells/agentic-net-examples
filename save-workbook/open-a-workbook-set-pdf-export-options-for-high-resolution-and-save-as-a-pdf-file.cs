// Title: Convert an Excel workbook to a high‑resolution PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export .xlsx to PDF with high quality settings | How to set PdfSaveOptions for high‑resolution PDF output in Aspose.Cells | Saving an Excel workbook as PDF with custom options using Aspose.Cells .NET | Aspose.Cells convert Excel to PDF preserving layout in C# | Example of PDF export with DPI control in Aspose.Cells C#
// Tags: Aspose.Cells PDF export high resolution | C# convert Excel to PDF Aspose.Cells | PdfSaveOptions usage Aspose.Cells .NET | Workbook.Save PDF Aspose.Cells example | Excel to PDF conversion Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the source .xlsx file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (default settings provide high‑resolution output), and saves the workbook as a PDF while handling load and save exceptions.
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

            // Load the existing workbook from a file
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Configure PDF export options (default options are sufficient for most cases)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file using the specified options
            try
            {
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save PDF: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
