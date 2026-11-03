// Title: Convert an Excel workbook to a single continuous PDF stream without Fast Web View using Aspose.Cells for .NET
// Common Searches: c# aspnet convert xlsx to pdf without fast web view | aspnet cells pdfsaveoptions how to produce single stream pdf | disable fast web view when exporting Excel to PDF with Aspose.Cells | save workbook as PDF continuous stream Aspose.Cells example | asp.net core generate PDF from Excel without incremental updates
// Tags: Aspose.Cells PdfSaveOptions continuous PDF output | Excel to PDF conversion Fast Web View disabled | C# produce single PDF file from workbook | PDF export incremental updates disabled Aspose.Cells | Aspose.Cells PDF export single-page stream

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies that the source Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (Fast Web View is not enabled by default), and saves the workbook as a single continuous PDF stream to the specified output path, with exception handling for any runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (Fast Web View is not supported in this version,
            // so we rely on the default behavior)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file with the specified options
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
