// Title: Refresh all PivotTables in an Excel workbook and export to PDF while preserving layout using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file, refreshes every PivotTable, recalculates all formulas, and saves the workbook as a PDF with the original layout using Aspose.Cells. | Generate a C# example that updates PivotTable data, applies PdfSaveOptions to keep document structure, and exports the refreshed workbook to a PDF file.
// Common Searches: c# aspose.cells refresh pivot tables before converting to pdf | preserve pivot table formatting when saving Excel as PDF with Aspose.Cells | export refreshed pivot table to PDF while keeping layout in .NET | how to use PdfSaveOptions ExportDocumentStructure for pivot tables
// Tags: refresh pivot tables Aspose.Cells | export workbook to PDF Aspose.Cells | preserve pivot table layout PDF | PdfSaveOptions ExportDocumentStructure | recalculate formulas before PDF export

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an Excel workbook, refreshes all PivotTables, recalculates formulas, and saves the workbook as a PDF while preserving the original layout and document structure.
class ExportPivotPdf
{
    static void Main()
    {
        // Load the workbook that contains the PivotTable
        Workbook workbook = new Workbook("input.xlsx");

        // Refresh all PivotTables to reflect the latest source data
        workbook.Worksheets.RefreshPivotTables();

        // Recalculate any formulas that might affect the PivotTable
        workbook.CalculateFormula();

        // Configure PDF save options to preserve the document structure (layout)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true
        };

        // Export the refreshed workbook (including the PivotTable) to PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
