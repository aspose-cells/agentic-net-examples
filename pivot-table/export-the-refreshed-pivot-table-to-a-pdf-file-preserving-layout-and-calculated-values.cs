using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Rendering;

class ExportPivotTablePdf
{
    static void Main()
    {
        // Load the workbook that contains the PivotTable(s)
        Workbook workbook = new Workbook("input.xlsx");

        // Refresh all PivotTables in the workbook to ensure calculated values are up‑to‑date
        workbook.Worksheets.RefreshPivotTables();

        // Configure PDF save options to preserve the document structure (layout, formatting, etc.)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true
        };

        // Save the refreshed workbook as a PDF file
        workbook.Save("refreshed_pivot.pdf", pdfOptions);
    }
}