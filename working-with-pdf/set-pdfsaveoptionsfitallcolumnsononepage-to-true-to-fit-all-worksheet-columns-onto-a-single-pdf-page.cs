// Title: C# – Fit All Worksheet Columns on a Single PDF Page with Aspose.Cells
// Description: This example shows how to create a workbook, populate it with many columns, enable the PdfSaveOptions.AllColumnsInOnePagePerSheet flag, and save the file as a PDF where each worksheet is compressed to fit on one page. Ideal for generating compact, printable PDFs from wide Excel sheets.
// Keywords: Aspose.Cells PDF export C# | AllColumnsInOnePagePerSheet | fit columns on one PDF page | single‑page Excel to PDF | C# workbook to PDF | compact PDF report Aspose | global PDF conversion
// Common Searches: Aspose.Cells fit all columns on one PDF page C# | PdfSaveOptions AllColumnsInOnePagePerSheet example | export wide Excel sheet to single‑page PDF | C# convert workbook to PDF with column fitting | how to compress Excel columns into one PDF page
// Developer Intent: Configure PdfSaveOptions so every worksheet’s columns are rendered on a single PDF page.
// Use Cases: Create email‑ready PDF reports from spreadsheets with dozens of columns. | Produce printable PDFs where each sheet automatically fits on one page, reducing paper usage. | Batch‑process multiple workbooks, applying the same single‑page layout to all sheets.
// AI Prompts: Generate C# code that uses Aspose.Cells to export a workbook to PDF with all columns on one page per sheet. | Explain the impact of the AllColumnsInOnePagePerSheet property on PDF layout and suggest complementary PdfSaveOptions. | Provide a step‑by‑step tutorial for configuring PdfSaveOptions to achieve single‑page column fitting in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFitAllColumnsDemo
{
    // This example shows how to create a workbook, populate it with many columns, enable the PdfSaveOptions.AllColumnsInOnePagePerSheet flag, and save the file as a PDF where each worksheet is compressed to fit on one page. Ideal for generating compact, printable PDFs from wide Excel sheets.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data spanning many columns
            for (int col = 0; col < 30; col++)
            {
                sheet.Cells[0, col].PutValue($"Header {col + 1}");
                sheet.Cells[1, col].PutValue($"Data {col + 1}");
            }

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Fit all columns of each sheet onto a single PDF page
            pdfOptions.AllColumnsInOnePagePerSheet = true;

            // Save the workbook as PDF using the configured options
            workbook.Save("FitAllColumnsOnOnePage.pdf", pdfOptions);
        }
    }
}
