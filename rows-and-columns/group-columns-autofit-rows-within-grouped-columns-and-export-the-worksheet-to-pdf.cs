// Title: C# – Group Columns, Auto‑Fit Rows, and Export Worksheet to Single‑Page PDF with Aspose.Cells
// Description: Demonstrates how to create a workbook, group a range of columns without hiding them, auto‑fit all rows to the grouped content, and save the sheet as a PDF where every column fits on one page using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | group columns Aspose.Cells | auto fit rows Aspose.Cells | PDF export Aspose.Cells | AllColumnsInOnePagePerSheet | OnePagePerSheet | Excel to PDF single page | .NET spreadsheet library | sample code | GitHub example
// Common Searches: Aspose.Cells group columns without hiding | auto fit rows after grouping columns C# | export Excel to single‑page PDF Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet example | C# code to group columns and fit rows
// Developer Intent: Group a specific column range, adjust row heights automatically, and generate a PDF that forces all columns onto one page.
// Use Cases: Produce printable reports with collapsible column sections while preserving full row visibility. | Create invoices or financial statements that fit on a single PDF page per sheet without manual scaling. | Export dashboards where column grouping and row auto‑fit keep the layout consistent in the final PDF.
// AI Prompts: Generate C# code using Aspose.Cells to group columns 2‑5, auto‑fit rows, and save the worksheet as a PDF with AllColumnsInOnePagePerSheet enabled. | Explain how OnePagePerSheet and AllColumnsInOnePagePerSheet affect PDF layout when exporting from Aspose.Cells. | Provide a step‑by‑step tutorial for grouping columns, auto‑fitting rows, and exporting to a single‑page PDF in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, group a range of columns without hiding them, auto‑fit all rows to the grouped content, and save the sheet as a PDF where every column fits on one page using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Fill the worksheet with sample data across several columns
        for (int row = 0; row < 10; row++)
        {
            for (int col = 0; col < 8; col++)
            {
                cells[row, col].PutValue($"Row{row + 1}_Col{col + 1} sample text");
            }
        }

        // Group columns 2 to 5 (zero‑based indices 1 to 4) without hiding them
        cells.GroupColumns(1, 4, false);

        // Auto‑fit all rows so that row heights adjust to the content in the grouped columns
        worksheet.AutoFitRows();

        // Configure PDF save options to place all columns on a single page per sheet
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            AllColumnsInOnePagePerSheet = true,
            OnePagePerSheet = true
        };

        // Export the worksheet to a PDF file
        workbook.Save("GroupedColumns_AutoFitRows.pdf", pdfOptions);
    }
}
