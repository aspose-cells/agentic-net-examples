using System;
using Aspose.Cells;

namespace AsposeCellsPivotPdfExport
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains the PivotTable(s)
            Workbook workbook = new Workbook("input.xlsx");

            // Refresh all PivotTables in the workbook to ensure calculated values are up‑to‑date
            workbook.Worksheets.RefreshPivotTables();

            // Configure PDF save options to preserve the document structure (layout)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Export the refreshed workbook (including the PivotTable layout) to a PDF file
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}