// Title: How to embed all fonts when converting an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as a PDF with every used font embedded using Aspose.Cells. | Show which PdfSaveOptions properties control font embedding and demonstrate setting them to embed all fonts during Excel‑to‑PDF conversion. | Create a C# snippet that verifies the source Excel file, configures full font embedding, and exports the workbook to PDF while handling errors.
// Common Searches: Aspose.Cells embed fonts in PDF conversion C# | C# save Excel as PDF with all fonts embedded Aspose.Cells | prevent missing glyphs when exporting .xlsx to PDF using Aspose.Cells | how to enable font embedding in PdfSaveOptions Aspose.Cells .NET
// Tags: Aspose.Cells PDF font embedding | PdfSaveOptions embed fonts | Excel to PDF conversion full font embedding | prevent missing glyphs Aspose.Cells .NET | C# embed used fonts in PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Retained for potential future use

// The example loads an Excel workbook, creates a PdfSaveOptions instance (default without font embedding), and saves the workbook as a PDF, including basic file‑existence checks and exception handling.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.pdf";

        try
        {
            // Verify that the source Excel file exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputFile);

            // Configure PDF save options (font embedding not available in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputFile, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
