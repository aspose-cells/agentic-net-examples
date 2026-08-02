// Title: C# – Group Columns, Auto‑Fit Rows, and Export to Single‑Page PDF with Aspose.Cells
// Description: Create a workbook, fill data, group columns A‑F, auto‑fit all rows, and save the sheet as a one‑page PDF using Aspose.Cells PdfSaveOptions.
// Keywords: Aspose.Cells group columns | auto fit rows Aspose.Cells | export to PDF Aspose.Cells | PdfSaveOptions OnePagePerSheet | C# Excel to PDF | column outlining PDF export
// Common Searches: Aspose.Cells group columns and export PDF | auto fit rows after column grouping C# | single page PDF from Excel Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet example | C# code to outline columns and create PDF
// Developer Intent: Group a specific column range, adjust row heights automatically, and generate a PDF where all columns appear on a single page.
// Use Cases: Produce printable reports with collapsed column sections while keeping the layout on one PDF page. | Create invoices where detail columns are grouped and rows are sized automatically before PDF conversion. | Distribute data grids as compact, single‑page PDFs for easy sharing and viewing.
// AI Prompts: Show C# code that groups columns A‑F, auto‑fits rows, and saves the worksheet as a one‑page PDF using Aspose.Cells. | Explain how to set PdfSaveOptions so that all columns are forced onto a single page after grouping. | Demonstrate combining column grouping, row auto‑fit, and PDF export in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsGroupedColumnsPdfExport
{
    // Create a workbook, fill data, group columns A‑F, auto‑fit all rows, and save the sheet as a one‑page PDF using Aspose.Cells PdfSaveOptions.
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // 2. Populate sample data across several columns and rows
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 6; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // 3. Group columns A (index 0) to F (index 5)
            //    This creates an outline for the specified column range
            cells.GroupColumns(0, 5);

            // 4. Auto‑fit all rows so that row heights adjust to the content
            //    (including the grouped columns)
            worksheet.AutoFitRows();

            // 5. Configure PDF save options to place all columns on a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,               // One page per sheet
                AllColumnsInOnePagePerSheet = true    // Force all columns onto that page
            };

            // 6. Save the workbook as a PDF file using the configured options
            workbook.Save("GroupedColumns_AutoFitRows.pdf", pdfOptions);

            Console.WriteLine("PDF exported successfully with grouped columns and auto‑fitted rows.");
        }
    }
}
