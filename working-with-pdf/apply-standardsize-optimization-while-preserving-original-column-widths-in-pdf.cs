// Title: Export Excel to PDF with Standard optimization and retain column widths using Aspose.Cells (C#)
// Description: Demonstrates how to set explicit column widths in a worksheet, configure PdfSaveOptions for Standard (high‑quality) optimization, and save the workbook as a PDF while preserving the original column layout.
// Keywords: Aspose.Cells PDF export | Standard optimization | preserve column width | PdfSaveOptions | C# Aspose.Cells | Excel to PDF column width | high quality PDF Aspose
// Common Searches: Aspose.Cells keep column width when saving to PDF | C# export workbook to PDF with Standard optimization | how to set column width before PDF conversion Aspose | standard quality PDF output Aspose.Cells
// Developer Intent: Create a PDF from an Excel workbook with high‑quality settings while ensuring the worksheet’s column widths remain unchanged.
// Use Cases: Producing printable reports where column alignment must match the original Excel layout. | Generating invoices or financial statements as PDFs that require exact column dimensions. | Distributing data tables to clients in PDF format without losing the designed column spacing.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to PDF with Standard optimization and keeps the defined column widths. | Show the steps to set custom column widths in Aspose.Cells before converting the sheet to a high‑quality PDF. | Explain how to configure PdfSaveOptions for Standard optimization while preserving column widths in the resulting PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to set explicit column widths in a worksheet, configure PdfSaveOptions for Standard (high‑quality) optimization, and save the workbook as a PDF while preserving the original column layout.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data
        worksheet.Cells["A1"].PutValue("Header");
        worksheet.Cells["A2"].PutValue("This is a longer piece of text that should retain its column width");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(12345);

        // Preserve original column widths by setting them explicitly.
        // (If you omit this step, Aspose.Cells will keep the default widths.)
        worksheet.Cells.SetColumnWidth(0, 20); // Column A width
        worksheet.Cells.SetColumnWidth(1, 15); // Column B width

        // Configure PDF save options to use Standard (high‑quality) optimization.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        pdfSaveOptions.OptimizationType = PdfOptimizationType.Standard; // High print quality

        // Save the workbook as a PDF while keeping the column widths unchanged.
        workbook.Save("StandardSizeOptimized.pdf", pdfSaveOptions);
    }
}
