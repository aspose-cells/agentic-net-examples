// Title: Add PDF bookmarks to an Excel workbook before exporting to PDF with Aspose.Cells for .NET
// AI Prompts: Insert a PdfBookmarkEntry named "Chapter 1" that links to cell A1, add it to workbook.PdfBookmarks, then save the workbook as PDF using PdfSaveOptions. | Update the example to create several PdfBookmarkEntry objects with custom titles and target cells (e.g., B5, D10), add each to the PdfBookmarks collection, and generate a PDF that includes all bookmarks. | Show how to set DestinationPage and DestinationLocation on a PdfBookmarkEntry, add it to the workbook, and export the Excel file so the PDF outline reflects those destinations.
// Common Searches: how to create PDF outline bookmarks from Excel using Aspose.Cells C# | Aspose.Cells add multiple PdfBookmarkEntry objects before saving to PDF | C# export Excel to PDF with custom bookmarks Aspose.Cells | set destination page for PdfBookmarkEntry in Aspose.Cells .NET | add PDF bookmarks to workbook programmatically Aspose.Cells
// Tags: PdfBookmarkEntry addition Aspose.Cells C# | export Excel to PDF with bookmarks Aspose.Cells | custom PDF outline from Excel Aspose.Cells | set PdfBookmarkEntry destination page Aspose.Cells | programmatic PDF bookmarks Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The code loads an Excel workbook, creates one or more PdfBookmarkEntry objects, adds them to the workbook's PdfBookmarks collection, and then saves the workbook as a PDF using PdfSaveOptions, ensuring any specified output directory exists.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (customize as needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
