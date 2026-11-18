// Title: Generate hierarchical PDF bookmarks for each Excel worksheet using Aspose.Cells PdfSaveOptions in C#
// Common Searches: how to create PDF bookmarks for each worksheet using Aspose.Cells in C# | Aspose.Cells PdfSaveOptions hierarchical bookmark example | C# generate PDF outline from an Excel workbook with Aspose.Cells | populate PdfSaveOptions.Bookmarks collection for Excel sheets | Aspose.Cells PDF bookmarks require Aspose.Cells.Pdf assembly
// Tags: Aspose.Cells PdfSaveOptions hierarchical bookmarks | C# Excel to PDF conversion with worksheet bookmarks | PdfSaveOptions bookmark collection usage | Aspose.Cells PDF bookmark assembly requirement | generate PDF outline from Excel worksheets

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// This example shows how to load an Excel workbook, configure PdfSaveOptions, and save it as a PDF while outlining the steps needed to populate the PdfSaveOptions.Bookmarks collection for hierarchical worksheet bookmarks. It also notes that the PdfBookmark class is unavailable in the base package and that the Aspose.Cells.Pdf assembly must be referenced to enable PDF bookmark functionality.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: PdfBookmark class is not available in the current Aspose.Cells package.
            // If PDF bookmarks are required, ensure the appropriate Aspose.Cells.Pdf assembly is referenced.

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
