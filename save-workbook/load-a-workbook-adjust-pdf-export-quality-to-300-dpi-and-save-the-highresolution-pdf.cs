// Title: Export Excel to High‑Resolution PDF (300 DPI) with Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx workbook using Aspose.Cells, configures PdfSaveOptions to resample images at 300 dpi with 100 % JPEG quality, and saves the result as a high‑resolution PDF.
// Keywords: Aspose.Cells | C# | PdfSaveOptions | SetImageResample | 300 DPI PDF | high resolution PDF | Excel to PDF export | image quality | Aspose.Cells .NET | PDF image resampling
// Common Searches: Aspose.Cells export PDF 300 dpi | PdfSaveOptions SetImageResample C# example | increase PDF image quality Aspose.Cells | save Excel as high resolution PDF .NET | Aspose.Cells high DPI PDF export
// Developer Intent: Generate a PDF from an Excel workbook with 300 dpi image resolution and maximum JPEG quality using Aspose.Cells for .NET.
// Use Cases: Print‑ready reports where graphics must stay sharp | Archival PDFs that preserve the exact visual fidelity of the original workbook | Marketing or presentation PDFs that require high‑resolution images
// AI Prompts: Provide a C# snippet that exports an Excel file to a 600 dpi PDF with 80 % JPEG quality using Aspose.Cells. | Show how to batch‑convert all .xlsx files in a folder to high‑resolution PDFs with custom DPI settings. | Explain how to embed fonts and adjust image resampling in PdfSaveOptions for optimal PDF output.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an .xlsx workbook using Aspose.Cells, configures PdfSaveOptions to resample images at 300 dpi with 100 % JPEG quality, and saves the result as a high‑resolution PDF.
class HighResolutionPdfExport
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options and set image resampling to 300 DPI with maximum JPEG quality
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.SetImageResample(300, 100); // 300 PPI, 100% JPEG quality

        // Save the workbook as a high‑resolution PDF
        workbook.Save("output_high_res.pdf", pdfOptions);
    }
}
