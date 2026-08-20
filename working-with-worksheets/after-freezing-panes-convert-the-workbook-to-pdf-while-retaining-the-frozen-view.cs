// Title: Freeze panes in Aspose.Cells and export to PDF while preserving the view (C#)
// Description: Demonstrates how to freeze specific rows and columns in an Aspose.Cells worksheet, enable ExportDocumentStructure, and save the workbook as a PDF so the frozen panes remain visible in the generated document.
// Keywords: Aspose.Cells FreezePanes PDF | C# export workbook to PDF with frozen panes | PdfSaveOptions ExportDocumentStructure | .NET Aspose.Cells freeze rows columns | preserve frozen view PDF conversion | Aspose.Cells worksheet freeze example
// Common Searches: keep frozen panes when converting Aspose.Cells workbook to PDF | Aspose.Cells C# export frozen rows and columns to PDF | PdfSaveOptions ExportDocumentStructure example | FreezePanes method PDF output Aspose.Cells | how to retain frozen view in PDF using Aspose.Cells
// Developer Intent: The developer needs to freeze selected rows/columns in a worksheet and then generate a PDF that maintains the frozen pane layout.
// Use Cases: Create printable reports where header rows stay fixed in the PDF. | Generate PDF dashboards from Excel files with frozen navigation panes. | Automate Excel‑to‑PDF conversion while preserving layout consistency for documentation.
// AI Prompts: Show how to change the frozen pane start cell to D4 and still export the PDF with the frozen view retained. | Add page orientation and margin settings to PdfSaveOptions while keeping ExportDocumentStructure enabled. | Explain how ExportDocumentStructure affects PDF output and when it might fail to preserve frozen panes.

using System;
using Aspose.Cells;

// Demonstrates how to freeze specific rows and columns in an Aspose.Cells worksheet, enable ExportDocumentStructure, and save the workbook as a PDF so the frozen panes remain visible in the generated document.
class FreezePaneToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        for (int row = 0; row < 20; row++)
        {
            sheet.Cells[row, 0].PutValue($"Row {row + 1}");
            sheet.Cells[row, 1].PutValue(row * 10);
        }

        // Freeze panes at cell C3 (row index 2, column index 2) with 2 rows and 2 columns frozen
        sheet.FreezePanes(2, 2, 2, 2);

        // Prepare PDF save options (optional settings)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Export document structure to keep pane information in the PDF
            ExportDocumentStructure = true
        };

        // Save the workbook as PDF; the frozen view is retained in the output
        workbook.Save("FrozenPaneOutput.pdf", pdfOptions);
    }
}
