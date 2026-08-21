// Title: Aspose.Cells C# – Separate PDF pages per worksheet using PdfSaveOptions.OnePagePerSheet
// Description: The example builds a workbook with two sheets, populates them with sample rows, enables the OnePagePerSheet flag in PdfSaveOptions, and saves the workbook as a PDF so that every sheet occupies its own page.
// Keywords: Aspose.Cells | PdfSaveOptions | OnePagePerSheet | C# PDF export | .NET workbook to PDF | multiple worksheets PDF | separate page per sheet | Aspose.Cells PDF options
// Common Searches: Aspose.Cells OnePagePerSheet C# example | save workbook as PDF each sheet on new page | PdfSaveOptions OnePagePerSheet usage | C# export multiple worksheets to PDF Aspose | force new page per worksheet Aspose.Cells
// Developer Intent: Enable the OnePagePerSheet flag in PdfSaveOptions so that each worksheet is rendered on an individual PDF page when the workbook is saved.
// Use Cases: Printing multi‑section reports where each section is a separate worksheet. | Generating individual invoice PDFs from a single workbook. | Creating a product catalog PDF with one sheet per page. | Exporting financial statements where each statement starts on a fresh page.
// AI Prompts: Provide a C# snippet that uses Aspose.Cells to save a workbook with two sheets as a PDF, placing each sheet on a separate page. | Explain how the OnePagePerSheet property affects PDF output in Aspose.Cells and how to configure it in .NET. | Show code that creates a multi‑sheet workbook, fills it with data, and exports it to PDF with a page break between worksheets using Aspose.Cells.

using System;
using Aspose.Cells;

// The example builds a workbook with two sheets, populates them with sample rows, enables the OnePagePerSheet flag in PdfSaveOptions, and saves the workbook as a PDF so that every sheet occupies its own page.
class Program
{
    static void Main()
    {
        // Create a new workbook and add two worksheets
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Populate each sheet with some sample data
        for (int i = 0; i < 20; i++)
        {
            sheet1.Cells[i, 0].PutValue($"Sheet1 Row {i + 1}");
            sheet2.Cells[i, 0].PutValue($"Sheet2 Row {i + 1}");
        }

        // Configure PDF save options to force each worksheet onto a separate page
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OnePagePerSheet = true; // Each sheet will be rendered on its own PDF page

        // Save the workbook as a PDF file using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
