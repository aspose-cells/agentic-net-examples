// Title: Export Visible Cells to PDF with Aspose.Cells for .NET (exclude hidden rows/columns)
// Description: Creates a workbook, hides selected rows and columns, uses ExportTableOptions (PlotVisibleRows, PlotVisibleColumns, PlotVisibleCells) to extract only visible data, copies that data into a new workbook, and saves it as a PDF. The result is a clean PDF report that omits any hidden rows or columns.
// Keywords: Aspose.Cells PDF export | exclude hidden rows Aspose | hide columns PDF .NET | ExportTableOptions visible cells | clean Excel PDF report | C# Aspose.Cells PDF conversion | visible data to PDF
// Common Searches: Aspose.Cells export PDF without hidden rows | C# hide rows columns then save as PDF | Export only visible cells to PDF using Aspose | Generate clean PDF report from Excel in .NET | PlotVisibleRows PlotVisibleColumns Aspose example
// Developer Intent: Produce a PDF that contains only the rows and columns currently visible in an Excel workbook.
// Use Cases: Financial statements where summary rows are hidden before PDF generation. | Printable reports that must not show user‑hidden helper columns. | Automated invoice PDFs created from Excel templates while omitting hidden calculation fields.
// AI Prompts: Show C# code to export only visible cells of an Aspose.Cells workbook to PDF. | How can I hide specific rows and columns and then save the remaining data as a PDF with Aspose.Cells? | Explain the role of ExportTableOptions PlotVisibleRows/Columns/Cells for creating a clean PDF report.

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// Creates a workbook, hides selected rows and columns, uses ExportTableOptions (PlotVisibleRows, PlotVisibleColumns, PlotVisibleCells) to extract only visible data, copies that data into a new workbook, and saves it as a PDF. The result is a clean PDF report that omits any hidden rows or columns.
class ExportPdfExcludingHidden
{
    static void Main()
    {
        try
        {
            // Create a workbook and fill it with sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            for (int r = 0; r < 10; r++)
            {
                for (int c = 0; c < 5; c++)
                {
                    cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                }
            }

            // Hide specific rows and columns
            worksheet.Cells.HideRow(2);   // hide row index 2 (third row)
            worksheet.Cells.HideRow(5);   // hide row index 5 (sixth row)
            worksheet.Cells.HideColumn(1); // hide column index 1 (second column)
            worksheet.Cells.HideColumn(3); // hide column index 3 (fourth column)

            // Export only the visible cells to a DataTable
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                PlotVisibleRows = true,
                PlotVisibleColumns = true,
                PlotVisibleCells = true
            };
            DataTable visibleData = worksheet.Cells.ExportDataTable(
                0, 0,
                worksheet.Cells.MaxDataRow + 1,
                worksheet.Cells.MaxDataColumn + 1,
                exportOptions);

            // Create a new workbook that contains only the visible data
            Workbook cleanWorkbook = new Workbook();
            Worksheet cleanSheet = cleanWorkbook.Worksheets[0];
            Cells cleanCells = cleanSheet.Cells;

            // Manually import the DataTable because ImportDataTable may not be available in all versions
            for (int i = 0; i < visibleData.Rows.Count; i++)
            {
                for (int j = 0; j < visibleData.Columns.Count; j++)
                {
                    cleanCells[i, j].PutValue(visibleData.Rows[i][j]);
                }
            }

            // Save the clean workbook to PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            string outputPath = "CleanReport.pdf";
            cleanWorkbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
