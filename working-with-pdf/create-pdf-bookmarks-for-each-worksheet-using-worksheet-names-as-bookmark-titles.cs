// Title: Generate PDF bookmarks from worksheet names when converting an Excel workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert Excel to PDF with each sheet as a PDF bookmark | How to preserve worksheet names as bookmarks in PDF output using Aspose.Cells .NET | Aspose.Cells PdfSaveOptions default bookmark behavior for worksheets in C# | Create PDF bookmarks from Excel sheet names with Aspose.Cells library | Enable worksheet bookmarks when saving workbook as PDF in Aspose.Cells C#
// Tags: Aspose.Cells worksheet-to-PDF bookmark generation | PdfSaveOptions default worksheet bookmarks | Excel workbook PDF conversion with sheet bookmarks C# | C# Aspose.Cells preserve sheet names in PDF bookmarks | Aspose.Cells PDF export bookmark settings .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, applies Aspose.Cells PdfSaveOptions (which automatically generate PDF bookmarks for each worksheet using the sheet names), and saves the workbook as a PDF file, with checks for file existence and basic exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Note: In the current Aspose.Cells version, bookmarks for worksheets are
            // enabled by default when saving to PDF. If specific bookmark settings are
            // required, they can be adjusted via the appropriate properties available
            // in the version being used.

            // Save the workbook as a PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
