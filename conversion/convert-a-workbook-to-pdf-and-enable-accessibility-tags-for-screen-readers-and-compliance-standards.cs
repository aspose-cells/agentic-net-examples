// Title: Convert an Excel workbook to PDF using Aspose.Cells for .NET – handling file checks and accessibility tag limitations
// Common Searches: how to save an .xlsx file as PDF with Aspose.Cells in C# | Aspose.Cells PDFSaveOptions example for .NET projects | C# code to convert Excel to PDF and verify input file existence | does Aspose.Cells support accessibility tags when exporting to PDF | error handling while saving a workbook to PDF with Aspose.Cells
// Tags: Aspose.Cells PDFSaveOptions usage | Excel to PDF conversion .NET | C# workbook.Save PDF example | accessibility tag limitation Aspose.Cells | input file existence check C#

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies that the source XLSX file exists, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions (noting that accessibility tags are not supported in the current version), and saves the workbook as a PDF while providing basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (accessible tags not supported in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
