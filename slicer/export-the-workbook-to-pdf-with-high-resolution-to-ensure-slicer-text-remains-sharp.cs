// Title: Export a Workbook to High‑Resolution PDF with Sharp Slicer Text using Aspose.Cells (C#)
// Description: Demonstrates how to set CellsHelper.DPI, configure PdfSaveOptions (Standard optimization, ExportDocumentStructure, image resample at 300 DPI, JPEG quality 100) and save a workbook as a PDF that preserves slicer label clarity.
// Keywords: Aspose.Cells PDF export C# | high DPI PDF Aspose | slicer text sharpness | PdfSaveOptions image resample | ExportDocumentStructure PDF | 300 DPI workbook export | high quality PDF generation | Aspose.Cells rendering settings
// Common Searches: Aspose.Cells export PDF 300 DPI | keep slicer labels clear in PDF | PdfSaveOptions high resolution example | how to set DPI for Aspose.Cells PDF | sharp slicer text Aspose.Cells C#
// Developer Intent: Generate a PDF from an Aspose.Cells workbook that retains crisp slicer labels by using high‑resolution rendering settings.
// Use Cases: Producing printable reports with slicers that meet corporate branding standards. | Creating dashboard PDFs for client presentations where label legibility is critical. | Archiving spreadsheets as PDFs for regulatory compliance while preserving visual fidelity of slicer controls.
// AI Prompts: Write C# code to export an Aspose.Cells workbook to a 600 DPI PDF while maintaining slicer formatting. | Explain the impact of ExportDocumentStructure on text rendering quality for slicers in a PDF. | Suggest a balance between file size and slicer sharpness by tweaking PdfSaveOptions optimization and image resample values.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to set CellsHelper.DPI, configure PdfSaveOptions (Standard optimization, ExportDocumentStructure, image resample at 300 DPI, JPEG quality 100) and save a workbook as a PDF that preserves slicer label clarity.
class ExportWorkbookToPdfHighRes
{
    static void Main()
    {
        // Set the DPI to a high value (e.g., 300) to improve rendering quality of slicer text.
        CellsHelper.DPI = 300;

        // Create a new workbook (or load an existing one).
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        // Add additional data / slicers as needed here.

        // Configure PDF save options for high‑resolution output.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use the standard optimization for maximum print quality.
            OptimizationType = PdfOptimizationType.Standard,
            // Export document structure (helps retain text clarity in PDF viewers).
            ExportDocumentStructure = true
        };

        // Resample images to 300 DPI with maximum JPEG quality (100).
        // This also influences the overall rendering DPI.
        pdfOptions.SetImageResample(300, 100);

        // Save the workbook as a high‑resolution PDF.
        workbook.Save("HighResOutput.pdf", pdfOptions);
    }
}
