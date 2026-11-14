// Title: How to save an Excel workbook to PDF while preserving slicer visual style using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook to PDF keep slicer formatting | embed slicer style in PDF with Aspose.Cells PdfSaveOptions | C# export Excel slicer to PDF preserving visual appearance | which Aspose.Cells version supports EmbedSlicerStyle property | error handling when saving Excel to PDF with slicer style in Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions embed slicer style | preserve slicer formatting PDF export .NET | Excel slicer visual style PDF conversion | Aspose.Cells version requirement EmbedSlicerStyle | C# exception handling workbook to PDF

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, creates a PdfSaveOptions object (noting that the EmbedSlicerStyle property is unavailable in the current Aspose.Cells version), and saves the workbook as a PDF while handling file‑not‑found and save‑time exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Configure PDF save options
        var pdfOptions = new PdfSaveOptions();
        // Note: The EmbedSlicerStyle property is not available in this version of Aspose.Cells.
        // If slicer style embedding is required, upgrade to a version that supports it.

        try
        {
            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving PDF: {ex.Message}");
        }
    }
}
