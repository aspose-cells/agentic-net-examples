// Title: Fit All Worksheet Columns onto One PDF Page with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to use Aspose.Cells PdfSaveOptions (OnePagePerSheet + AllColumnsInOnePagePerSheet) in C# to convert an Excel workbook so that every column fits on a single PDF page per sheet.
// Keywords: Aspose.Cells | PdfSaveOptions | AllColumnsInOnePagePerSheet | OnePagePerSheet | C# | .NET | Excel to PDF | fit columns | single page PDF | wide spreadsheet export
// Common Searches: Aspose.Cells fit all columns on one PDF page | C# export Excel to single‑page PDF using PdfSaveOptions | OnePagePerSheet example Aspose.Cells | AllColumnsInOnePagePerSheet usage guide | scale columns when converting Excel to PDF with Aspose.Cells
// Developer Intent: Generate a PDF where each worksheet’s columns are rendered on a single page.
// Use Cases: Produce compact, printable PDFs for wide spreadsheets without manual scaling. | Automate one‑page PDF reports for financial statements or dashboards. | Create email‑ready PDFs of large data tables that must fit on a single sheet.
// AI Prompts: Show how to adjust column widths before applying AllColumnsInOnePagePerSheet in Aspose.Cells. | Provide code to set page orientation, margins, and scaling while using OnePagePerSheet. | Explain how to combine AllColumnsInOnePagePerSheet with custom paper size for optimal fit.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfFitColumns
{
    // Demonstrates how to use Aspose.Cells PdfSaveOptions (OnePagePerSheet + AllColumnsInOnePagePerSheet) in C# to convert an Excel workbook so that every column fits on a single PDF page per sheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with many columns to demonstrate fitting
            for (int col = 0; col < 50; col++)
            {
                // Header
                sheet.Cells[0, col].PutValue($"Column {col + 1}");
                // Sample data
                sheet.Cells[1, col].PutValue($"Data {col + 1}");
            }

            // Configure PDF save options to fit all columns on a single page per sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,               // Ensure each sheet is rendered as one page
                AllColumnsInOnePagePerSheet = true    // Fit all columns onto that page
            };

            // Save the workbook as PDF using the configured options
            workbook.Save("AllColumnsOnePage.pdf", pdfOptions);
        }
    }
}
