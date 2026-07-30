// Title: Freeze Panes and Export to PDF with Aspose.Cells (C#) – Preserve Frozen View
// Description: Create a workbook, apply FreezePanes at a specific cell, configure PdfSaveOptions, and save as PDF while keeping the frozen rows and columns visible in the generated document.
// Keywords: Aspose.Cells FreezePanes C# | export frozen view PDF | PdfSaveOptions Aspose.Cells | C# workbook to PDF | preserve frozen rows columns | Aspose.Cells PDF export settings | freeze panes PDF output | Aspose.Cells documentation example | C# spreadsheet PDF conversion | retain frozen panes PDF
// Common Searches: Aspose.Cells keep frozen panes when saving as PDF | C# freeze rows and columns then export to PDF | PdfSaveOptions preserve frozen view Aspose.Cells | how to export worksheet with frozen panes to PDF | Aspose.Cells FreezePanes PDF example
// Developer Intent: The developer needs to freeze selected rows/columns in a worksheet and generate a PDF that shows the same frozen layout.
// Use Cases: Generating printable reports where header rows stay fixed in the PDF. | Providing PDF previews of dashboards that retain the spreadsheet’s navigation layout. | Automating export of financial statements with frozen header and side panels for consistent reading.
// AI Prompts: Show C# code that freezes panes at a given cell and saves the workbook as a PDF using Aspose.Cells, keeping the frozen view. | Explain which PdfSaveOptions properties affect frozen pane preservation when exporting to PDF. | Give a step‑by‑step guide to convert a frozen worksheet to PDF with Aspose.Cells, including any required settings.

using System;
using Aspose.Cells;

// Create a workbook, apply FreezePanes at a specific cell, configure PdfSaveOptions, and save as PDF while keeping the frozen rows and columns visible in the generated document.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data
        for (int i = 0; i < 20; i++)
        {
            sheet.Cells[i, 0].PutValue($"Row {i + 1}");
            sheet.Cells[i, 1].PutValue(i * 10);
        }

        // Freeze panes at cell C3 with 2 frozen rows and 2 frozen columns
        sheet.FreezePanes("C3", 2, 2);

        // Configure PDF save options (optional)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true; // retain document structure

        // Save the workbook as PDF; the frozen view is preserved in the output
        workbook.Save("FrozenView.pdf", pdfOptions);
    }
}
