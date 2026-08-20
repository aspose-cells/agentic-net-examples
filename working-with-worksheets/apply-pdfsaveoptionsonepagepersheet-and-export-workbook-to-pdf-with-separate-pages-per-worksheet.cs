// Title: Export Workbook to PDF with One Page per Worksheet using Aspose.Cells (C#)
// Description: Shows how to build a workbook with multiple sheets, fill them with data, enable PdfSaveOptions.OnePagePerSheet, and save the file as a single PDF where each worksheet appears on a separate page.
// Keywords: Aspose.Cells | PdfSaveOptions | OnePagePerSheet | C# PDF export | worksheet to PDF | single PDF per sheet | Aspose.Cells PDF options | export multiple sheets
// Common Searches: Aspose.Cells export each worksheet to separate PDF page | PdfSaveOptions OnePagePerSheet C# example | save workbook as PDF one page per sheet | C# Aspose.Cells PDF per sheet | how to generate PDF with one page per worksheet Aspose
// Developer Intent: Create a PDF where every worksheet is rendered on its own page.
// Use Cases: Produce a printable report that consolidates several worksheets into a single PDF, with one page per sheet. | Generate a batch of invoices where each worksheet represents an invoice and the final PDF contains one page per invoice. | Share data‑analysis workbooks with stakeholders, preserving a clear one‑page‑per‑sheet layout for quick review.
// AI Prompts: Provide C# code that uses Aspose.Cells to export a multi‑sheet workbook to a PDF with OnePagePerSheet enabled. | Explain the impact of setting PdfSaveOptions.OnePagePerSheet to true versus false in Aspose.Cells. | Show an example that creates three worksheets, adds sample data, and saves them as a single PDF with each sheet on a separate page.

using System;
using Aspose.Cells;

// Shows how to build a workbook with multiple sheets, fill them with data, enable PdfSaveOptions.OnePagePerSheet, and save the file as a single PDF where each worksheet appears on a separate page.
class Program
{
    static void Main()
    {
        // Create a new workbook (default workbook contains one worksheet)
        Workbook workbook = new Workbook();

        // Reference the first worksheet and give it a name
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Add a second worksheet
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Fill both worksheets with sample data
        for (int i = 0; i < 20; i++)
        {
            sheet1.Cells[i, 0].PutValue($"Sheet1 Row {i + 1}");
            sheet2.Cells[i, 0].PutValue($"Sheet2 Row {i + 1}");
        }

        // Create PDF save options and enable OnePagePerSheet
        // This forces each worksheet to be rendered on a single PDF page
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true
        };

        // Save the workbook to a PDF file using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
