// Title: C# – Remove Printer Settings, Set A2 Paper Size, and Export to PDF with Aspose.Cells
// Description: Demonstrates how to clear a worksheet's printer settings, change the page format to A2, and save the workbook as a size‑optimized PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# PDF export | clear printer settings Aspose.Cells | A2 paper size Excel | PdfSaveOptions minimum size | worksheet PageSetup printer settings
// Common Searches: how to clear printer settings before PDF export Aspose.Cells | set A2 page size for Excel PDF using Aspose.Cells | optimize PDF file size with Aspose.Cells .NET | remove printer metadata from Excel PDF | C# Aspose.Cells export A2 PDF
// Developer Intent: The developer needs to strip printer configuration, apply an A2 page layout, and generate a compact PDF from a workbook.
// Use Cases: Archiving Excel reports without embedding printer-specific data. | Producing A2‑sized PDFs for posters, schematics, or large‑format prints. | Delivering the smallest possible PDF for web download or email attachment.
// AI Prompts: Provide C# code that clears a worksheet's printer settings, sets the page size to A2, and saves the workbook as a minimal‑size PDF with Aspose.Cells. | Show an example of using PdfSaveOptions to reduce PDF file size after changing PageSetup in Aspose.Cells. | Explain why removing printer settings before PDF conversion can prevent unwanted metadata in the output file.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to clear a worksheet's printer settings, change the page format to A2, and save the workbook as a size‑optimized PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (optional: add some data)
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Demo for PDF export");

        // Remove any printer settings by clearing the byte array
        worksheet.PageSetup.PrinterSettings = null;

        // Set the paper size to A2
        worksheet.PageSetup.PaperSize = PaperSizeType.PaperA2;

        // Prepare PDF save options (e.g., minimize file size)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}
