// Title: Convert an Excel workbook to PDF with system default font substitution using Aspose.Cells for .NET
// Common Searches: how to export an .xlsx file to PDF with fallback to Arial using Aspose.Cells in C# | Aspose.Cells PDFSaveOptions DefaultFont property to handle missing fonts in .NET | C# code to convert Excel to PDF and replace unavailable fonts with a system default font | save workbook as PDF with font substitution when original fonts are not installed Aspose.Cells | using Aspose.Cells to generate PDF from Excel with default font fallback for missing fonts
// Tags: Aspose.Cells PDFSaveOptions DefaultFont | Excel to PDF conversion with font fallback | C# missing font substitution Aspose.Cells | PDF export using system default font | Workbook.Save PDF default font option

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the source XLSX file, loads it into an Aspose.Cells Workbook, sets PdfSaveOptions.DefaultFont to a system font (e.g., Arial) to replace any missing fonts, and saves the workbook as a PDF, including basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to use a default system font for missing fonts
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Specify a common system font to replace missing fonts (e.g., Arial)
                DefaultFont = "Arial"
            };

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
