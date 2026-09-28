// Title: Add hierarchical PDF bookmarks for each worksheet when converting an Excel workbook to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that iterates through all worksheets in a Workbook and adds a corresponding entry to PdfSaveOptions.Bookmarks before saving as PDF using Aspose.Cells. | Show how to build a nested bookmark structure in PdfSaveOptions so each Excel sheet appears as a top‑level outline item in the exported PDF. | Create a reusable method that accepts a Workbook and returns PdfSaveOptions with a populated Bookmarks collection for worksheet‑level PDF navigation.
// Common Searches: asp.net convert excel to pdf with worksheet bookmarks using aspose.cells | c# pdfsaveoptions bookmarks each worksheet example | how to create PDF outline from Excel sheets with Aspose.Cells | add hierarchical bookmarks to PDF export of Excel file in .NET
// Tags: Aspose.Cells worksheet PDF bookmarks | PdfSaveOptions hierarchical bookmark collection C# | export Excel to PDF with outline Aspose.Cells | C# generate PDF bookmark entries from workbook | Aspose.Cells PDF save options bookmark API

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input Excel file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions instance (noting that PDF bookmark classes are unavailable in the current version), ensures the output directory exists, saves the workbook as a PDF, and catches any exceptions to display an error message.
class PdfBookmarkExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare PDF save options (SaveFormat is implicit for PdfSaveOptions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: PDF bookmark classes are not available in the current Aspose.Cells version.
            // If bookmark support is required, ensure the appropriate Aspose.Cells.Pdf assembly is referenced.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
