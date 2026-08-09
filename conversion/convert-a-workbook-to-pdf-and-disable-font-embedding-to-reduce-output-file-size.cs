// Title: C# Aspose.Cells – Convert Excel to PDF without Font Embedding for Minimal File Size
// Description: Demonstrates how to use Aspose.Cells for .NET to save an Excel workbook as a PDF while disabling standard Windows font embedding and applying the MinimumSize optimization, resulting in a lightweight PDF suitable for fast distribution.
// Keywords: Aspose.Cells PDF conversion C# | disable font embedding Aspose.Cells | minimum size PDF Aspose.Cells | PdfSaveOptions EmbedStandardWindowsFonts false | Excel to PDF small file size | C# reduce PDF size Aspose.Cells
// Common Searches: Aspose.Cells disable font embedding when saving PDF | C# convert Excel to PDF with smallest size | PdfSaveOptions MinimumSize Aspose.Cells example | How to create lightweight PDF from workbook using Aspose.Cells | C# Aspose.Cells PDF optimization settings
// Developer Intent: Create a PDF from an Excel workbook while turning off font embedding to keep the output file size as low as possible.
// Use Cases: Generating compact PDF reports for email or web preview when target devices already have standard Windows fonts. | Batch‑processing large numbers of spreadsheets into small PDFs for archival or cloud storage. | Automating invoice or statement PDFs that need to be lightweight for fast download.
// AI Prompts: Show a C# example that converts an .xlsx to PDF with Aspose.Cells, disables font embedding, and uses MinimumSize optimization. | Explain how the EmbedStandardWindowsFonts and OptimizationType properties affect PDF size in Aspose.Cells. | Provide step‑by‑step code to load an existing workbook, set PdfSaveOptions to not embed fonts, and save a reduced‑size PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to use Aspose.Cells for .NET to save an Excel workbook as a PDF while disabling standard Windows font embedding and applying the MinimumSize optimization, resulting in a lightweight PDF suitable for fast distribution.
class ConvertWorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new Workbook("input.xlsx") to load

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample text for PDF conversion.");

        // Configure PDF save options to disable font embedding and minimize file size
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.EmbedStandardWindowsFonts = false;               // Disable embedding of standard Windows fonts
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize; // Optimize for smallest file size

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
