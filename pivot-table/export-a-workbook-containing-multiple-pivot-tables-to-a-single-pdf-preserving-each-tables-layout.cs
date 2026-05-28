using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class ExportPivotTablesToPdf
{
    static void Main()
    {
        // Load the workbook that contains multiple pivot tables
        Workbook workbook = new Workbook("input.xlsx");

        // Refresh all pivot tables in each worksheet to ensure they reflect the latest data
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.RefreshPivotTables();
        }

        // Alternatively, refresh all pivot tables in the workbook at once
        workbook.Worksheets.RefreshPivotTables();

        // Configure PDF save options to preserve the document structure (keeps pivot table layouts)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true,
            // Optional: keep each worksheet on separate pages
            OnePagePerSheet = false
        };

        // Save the entire workbook (all worksheets) as a single PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}