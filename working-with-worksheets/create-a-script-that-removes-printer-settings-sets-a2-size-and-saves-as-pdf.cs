// Title: C# – Remove Printer Settings, Set A2 Paper Size, and Export to PDF with Aspose.Cells
// Description: Creates a new Workbook, clears any printer settings on the first worksheet, sets the workbook's default paper size to A2, applies PdfSaveOptions for minimum file size, and saves the result as a PDF.
// Keywords: Aspose.Cells C# remove printer settings | Aspose.Cells set A2 paper size | Aspose.Cells export PDF | PdfSaveOptions MinimumSize | C# workbook to PDF | Aspose.Cells page setup | A2 format Aspose.Cells
// Common Searches: clear printer settings Aspose.Cells before PDF export | set A2 page size Aspose.Cells .NET | export workbook to PDF with smallest file size Aspose.Cells | how to change paper size for all worksheets Aspose.Cells | remove page margins Aspose.Cells PDF conversion
// Developer Intent: Remove existing printer configuration, apply an A2 page layout, and generate a compact PDF from a workbook using Aspose.Cells for .NET.
// Use Cases: Eliminate unwanted printer margins when converting spreadsheets to PDF. | Produce large‑format (A2) reports or posters directly from code. | Reduce PDF file size for faster web delivery or email attachment.
// AI Prompts: Provide C# code that clears printer settings, sets A2 paper size, and saves a workbook as a minimal‑size PDF using Aspose.Cells. | Explain the effect of PdfOptimizationType.MinimumSize on PDF output in Aspose.Cells. | Show how to set a global paper size for a workbook and its impact on PDF rendering in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // Creates a new Workbook, clears any printer settings on the first worksheet, sets the workbook's default paper size to A2, applies PdfSaveOptions for minimum file size, and saves the result as a PDF.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Remove any printer settings (set to null)
            worksheet.PageSetup.PrinterSettings = null;

            // Set the default paper size to A2 (420 x 594 mm)
            workbook.Settings.PaperSize = PaperSizeType.PaperA2;

            // Prepare PDF save options (optional: minimize file size)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

            // Save the workbook as a PDF file
            workbook.Save("Output_A2.pdf", pdfOptions);
        }
    }
}
