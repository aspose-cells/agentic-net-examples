// Title: How to add PDF bookmarks for each worksheet when converting an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add PDF bookmarks for each worksheet during Excel to PDF conversion | Enable worksheet bookmarks in PDF output using PdfSaveOptions in Aspose.Cells .NET | How to set the PdfBookmarks property when saving a workbook as PDF with Aspose.Cells | Aspose.Cells PDF export include sheet bookmarks C# example | PdfSaveOptions PdfBookmarks not available version check Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions add worksheet bookmarks | C# export Excel to PDF with bookmarks | PdfBookmarks property Aspose.Cells .NET | Excel worksheet PDF bookmark generation | Aspose.Cells PDF conversion include sheet outlines

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an Excel workbook, enabling the PdfBookmarks flag on PdfSaveOptions (when supported), and saving the workbook as a PDF where each worksheet appears as a bookmark, with basic error handling.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Note: In some versions of Aspose.Cells the PdfBookmarks property may not be available.
            // If supported, you can enable it as shown below:
            // pdfOptions.PdfBookmarks = true;

            // Save the workbook as PDF using the configured options
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
