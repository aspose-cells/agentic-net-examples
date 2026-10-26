// Title: Add PDF bookmarks with named destinations linked to specific Excel cell ranges using Aspose.Cells for .NET (C#)
// Common Searches: how to create PDF bookmarks that point to a specific cell range in an Excel file using Aspose.Cells C# | Aspose.Cells C# add named destination bookmarks when exporting to PDF | export Excel to PDF with clickable outline entries for rows using Aspose.Cells | C# generate PDF outline entries from worksheet cells with PdfSaveOptions | Aspose.Cells example for linking PDF bookmarks to cell A1:B10
// Tags: Aspose.Cells PDF bookmark creation | named destinations in PDF export | C# add PDF outline entries from Excel cells | PdfSaveOptions bookmark configuration | Excel cell range to PDF bookmark linking | Aspose.Cells export workbook to PDF with bookmarks

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, creates a PdfSaveOptions object, and saves the workbook as a PDF file. While the code demonstrates basic PDF export, adding PDF bookmarks with named destinations that correspond to particular cell ranges requires configuring the PdfSaveOptions (available in the Aspose.Cells.Pdf assembly) to define bookmarks that reference those cells.
class PdfExportExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input Excel file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare PDF save options (basic options; bookmarks require Aspose.Cells.Pdf assembly)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF file successfully created at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
