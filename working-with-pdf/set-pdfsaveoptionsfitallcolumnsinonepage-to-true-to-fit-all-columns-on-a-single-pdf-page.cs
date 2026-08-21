// Title: Export Excel to a Single‑Page PDF with All Columns Visible – Aspose.Cells C# Example
// Description: Shows how to build a workbook, fill it with many columns, and apply Aspose.Cells PdfSaveOptions (AllColumnsInOnePagePerSheet = true) to produce a PDF where each worksheet’s columns are compressed onto one page.
// Keywords: Aspose.Cells PDF export C# | PdfSaveOptions AllColumnsInOnePagePerSheet | fit all columns on one PDF page | .NET Excel to PDF single page | export wide worksheet to PDF | C# Aspose.Cells example PDF | single‑page PDF from Excel
// Common Searches: Aspose.Cells set AllColumnsInOnePagePerSheet true | C# export Excel to one‑page PDF | fit all columns in one PDF page Aspose.Cells | PdfSaveOptions single page per sheet example | how to compress Excel columns into one PDF page .NET
// Developer Intent: Configure PdfSaveOptions so that every column of each worksheet is rendered on a single PDF page during the save operation.
// Use Cases: Create compact PDF reports from wide spreadsheets for email attachment. | Generate printable one‑page dashboards that contain many data columns. | Produce consistent, single‑page PDFs for financial statements or invoices with numerous columns.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to PDF with AllColumnsInOnePagePerSheet enabled. | Explain the difference between AllColumnsInOnePagePerSheet and OnePagePerSheet in Aspose.Cells PDF conversion. | Provide a snippet that sets PdfSaveOptions to fit all columns on one page and also customizes page margins for optimal layout.

using System;
using Aspose.Cells;

namespace AsposeCellsFitAllColumnsDemo
{
    // Shows how to build a workbook, fill it with many columns, and apply Aspose.Cells PdfSaveOptions (AllColumnsInOnePagePerSheet = true) to produce a PDF where each worksheet’s columns are compressed onto one page.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data spanning many columns
            for (int col = 0; col < 30; col++)
            {
                // Header row
                sheet.Cells[0, col].PutValue($"Header {col + 1}");
                // Data row
                sheet.Cells[1, col].PutValue($"Data {col + 1}");
            }

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Fit all columns of each sheet onto a single PDF page
            pdfOptions.AllColumnsInOnePagePerSheet = true;

            // Optionally, you can also set OnePagePerSheet if you want the entire sheet on one page
            // pdfOptions.OnePagePerSheet = true;

            // Save the workbook as PDF using the configured options
            workbook.Save("AllColumnsOnePage.pdf", pdfOptions);

            Console.WriteLine("PDF saved with all columns fitted on a single page.");
        }
    }
}
