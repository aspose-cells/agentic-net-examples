// Title: Convert an Excel workbook to PDF using Aspose.Cells for .NET with PdfSaveOptions
// Common Searches: how to save an Excel workbook as PDF with Aspose.Cells in C# | Aspose.Cells PdfSaveOptions example for converting .xlsx to .pdf | C# check file existence before converting Excel to PDF using Aspose.Cells | set default PDF export settings with Aspose.Cells Workbook.Save | log conversion success or errors when exporting Excel to PDF in .NET
// Tags: Aspose.Cells PDF conversion C# | PdfSaveOptions usage Aspose.Cells | Excel to PDF export .NET | Workbook.Save PDF Aspose.Cells | file existence validation C# Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object with default settings, and saves the workbook as a PDF to the specified output path, outputting success or error messages to the console.
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

            // Configure PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
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
