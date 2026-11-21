// Title: How to preserve Excel workbook bookmarks as PDF outline entries using ExportDocumentStructure in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel bookmarks to PDF outline | Enable ExportDocumentStructure to keep Excel named ranges as PDF bookmarks | Convert .xlsx to PDF with outline entries using Aspose.Cells PdfSaveOptions | C# save workbook as PDF while preserving document structure | Aspose.Cells PDF save options retain Excel bookmark hierarchy
// Tags: ExportDocumentStructure PDF outline Aspose.Cells | Excel bookmarks PDF conversion C# | PdfSaveOptions document structure preservation | named ranges as PDF bookmarks Aspose.Cells | C# workbook to PDF with outline entries

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures PdfSaveOptions.ExportDocumentStructure to true to generate PDF outline entries that reflect Excel bookmarks, and saves the workbook as a PDF while handling missing files and exceptions.
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

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to retain document structure (outline)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true   // Enables PDF outline entries
                // Note: ExportBookmarks property is not available in current Aspose.Cells versions.
            };

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF file saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
