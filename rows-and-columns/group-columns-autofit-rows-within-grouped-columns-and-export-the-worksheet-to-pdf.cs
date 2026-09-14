// Title: How to group columns, auto‑fit rows, and save a worksheet as a single‑page PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Group columns B through D, hide them, auto‑fit rows 0‑9, and export the worksheet to a PDF where all columns fit on one page using Aspose.Cells in C#. | Create a workbook, fill it with data, apply column grouping, auto‑fit rows and columns, then save as a PDF with OnePagePerSheet and AllColumnsInOnePagePerSheet options via Aspose.Cells.
// Common Searches: Aspose.Cells C# group a range of columns and hide them | auto fit rows after grouping columns in Aspose.Cells .NET | save Aspose.Cells worksheet to PDF with all columns on one page | PdfSaveOptions OnePagePerSheet and AllColumnsInOnePagePerSheet example C# | how to auto‑fit rows for a specific range in Aspose.Cells
// Tags: group columns hide Aspose.Cells C# | auto‑fit rows after column grouping Aspose.Cells | export worksheet to PDF single page Aspose.Cells | PdfSaveOptions OnePagePerSheet Aspose.Cells | auto‑fit columns Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsGroupAndPdfDemo
{
    // The example creates a workbook, populates sample data, groups and hides columns B‑D, auto‑fits rows 0‑9 and all columns, and saves the sheet as a PDF where every column fits on a single page per sheet using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data across several columns (A to E)
            for (int row = 0; row < 10; row++)
            {
                cells[row, 0].PutValue($"Item {row + 1}");
                cells[row, 1].PutValue($"Description for item {row + 1} that may be quite long");
                cells[row, 2].PutValue(row * 10);
                cells[row, 3].PutValue(DateTime.Now.AddDays(row).ToShortDateString());
                cells[row, 4].PutValue($"Notes {row + 1}");
            }

            // Group columns B (index 1) to D (index 3) and hide them
            cells.GroupColumns(1, 3, true);

            // Auto‑fit rows that contain data within the grouped columns.
            // Since rows 0‑9 contain data, we autofit those rows.
            worksheet.AutoFitRows(0, 9);

            // Optionally auto‑fit all columns for better visibility
            worksheet.AutoFitColumns();

            // Prepare PDF save options: fit all columns on one page per sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true,
                AllColumnsInOnePagePerSheet = true,
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF
            workbook.Save("GroupedColumnsAutoFitRows.pdf", pdfOptions);
        }
    }
}
