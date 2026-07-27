// Title: Export Excel to PDF with Aspose.Cells – omit hidden rows & columns (C#)
// Description: Demonstrates how to create a workbook, hide specific rows and columns, and save it as a PDF using Aspose.Cells. Hidden rows and columns are automatically excluded from the PDF output, delivering a clean report without extra configuration.
// Keywords: Aspose.Cells PDF export C# | hide rows PDF Aspose | exclude hidden columns PDF | clean Excel PDF report | PdfSaveOptions hidden rows | C# Aspose.Cells generate PDF | Excel to PDF without hidden cells
// Common Searches: Aspose.Cells export PDF hidden rows C# | C# hide column then save as PDF Aspose | skip hidden rows when converting Excel to PDF using Aspose | generate PDF report from Excel ignoring hidden cells | PdfSaveOptions hide rows Aspose.Cells
// Developer Intent: Create a PDF from an Excel workbook that includes only the visible rows and columns.
// Use Cases: Produce a financial summary PDF where calculation rows are hidden before export. | Generate printable invoices that omit helper columns and rows for a tidy layout. | Automate batch conversion of multiple Excel files to PDFs while ensuring hidden data is never rendered.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to PDF, guaranteeing hidden rows and columns are not rendered. | Show how to use PdfSaveOptions to control page handling while confirming hidden cells are automatically excluded. | Explain how to programmatically verify that hidden rows and columns are omitted from the generated PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for SheetSet if needed

// Demonstrates how to create a workbook, hide specific rows and columns, and save it as a PDF using Aspose.Cells. Hidden rows and columns are automatically excluded from the PDF output, delivering a clean report without extra configuration.
class ExportPdfExcludingHidden
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Fill the worksheet with sample data
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["C1"].PutValue("Header3");
        worksheet.Cells["A2"].PutValue("Data1");
        worksheet.Cells["B2"].PutValue("Data2");
        worksheet.Cells["C2"].PutValue("Data3");
        worksheet.Cells["A3"].PutValue("Data4");
        worksheet.Cells["B3"].PutValue("Data5");
        worksheet.Cells["C3"].PutValue("Data6");

        // Hide a row (row index 1 -> second row) and a column (column index 1 -> column B)
        worksheet.Cells.HideRow(1);
        worksheet.Cells.HideColumn(1);

        // Configure PDF save options.
        // Hidden rows and columns are automatically omitted during PDF rendering,
        // so no special option is required beyond the standard save.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Optional: control page handling (default prints all pages)
            PrintingPageType = PrintingPageType.Default
        };

        // Save the workbook to PDF. The resulting file will contain only the visible rows and columns.
        workbook.Save("CleanReport.pdf", pdfOptions);
    }
}
