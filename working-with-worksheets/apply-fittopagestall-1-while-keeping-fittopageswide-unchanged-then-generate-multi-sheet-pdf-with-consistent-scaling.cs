// Title: Aspose.Cells .NET – Set FitToPagesTall = 1, preserve FitToPagesWide, export all worksheets to a single PDF
// Description: Creates a workbook with two worksheets, fills them with sample data, disables percent scaling, keeps the existing FitToPagesWide, sets FitToPagesTall to 1, defines a dynamic print area for each sheet, and saves the file as a multi‑sheet PDF using PdfSaveOptions that respect the vertical fit‑to‑page setting.
// Keywords: Aspose.Cells | FitToPagesTall | FitToPagesWide | C# PDF export | multi‑sheet PDF | PageSetup scaling | dynamic print area | PdfSaveOptions | .NET Excel to PDF | vertical fit to page
// Common Searches: Aspose.Cells set FitToPagesTall 1 without changing FitToPagesWide | export multiple worksheets to one PDF in C# | how to keep horizontal scaling when fitting to one page tall | define print area programmatically before PDF conversion Aspose.Cells | PdfSaveOptions allow multiple pages per sheet Aspose.Cells
// Developer Intent: Apply a one‑page‑tall vertical scaling to every worksheet while leaving the horizontal scaling unchanged, then generate a combined PDF of all sheets.
// Use Cases: Produce a printable report where each sheet must fit vertically on a single page regardless of column width. | Create invoices or statements with varying column counts that need a consistent height across a multi‑sheet PDF. | Automate the conversion of large data tables across several worksheets to PDF while preserving vertical page layout.
// AI Prompts: Show C# code to set FitToPagesTall = 1 for all worksheets in an Aspose.Cells workbook while keeping FitToPagesWide unchanged and export to a single PDF. | How can I dynamically set the print area for each worksheet based on its used range before saving as PDF with Aspose.Cells? | Explain PdfSaveOptions settings that allow multiple pages per sheet but still respect FitToPagesTall = 1 in Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsFitToPagesTallDemo
{
    // Creates a workbook with two worksheets, fills them with sample data, disables percent scaling, keeps the existing FitToPagesWide, sets FitToPagesTall to 1, defines a dynamic print area for each sheet, and saves the file as a multi‑sheet PDF using PdfSaveOptions that respect the vertical fit‑to‑page setting.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sheet 1 – populate with sample data
            // -------------------------------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            for (int row = 0; row < 120; row++)          // enough rows to span several pages
            {
                for (int col = 0; col < 8; col++)
                {
                    sheet1.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // -------------------------------------------------
            // Sheet 2 – populate with different sample data
            // -------------------------------------------------
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            for (int row = 0; row < 80; row++)
            {
                for (int col = 0; col < 12; col++)
                {
                    sheet2.Cells[row, col].PutValue($"S{row + 1}C{col + 1}");
                }
            }

            // -------------------------------------------------
            // Apply page‑setup scaling: FitToPagesTall = 1,
            // keep FitToPagesWide unchanged (default = 1)
            // -------------------------------------------------
            foreach (Worksheet ws in workbook.Worksheets)
            {
                PageSetup ps = ws.PageSetup;

                // Ensure scaling is driven by FitToPagesWide/Tall, not by percent zoom
                ps.IsPercentScale = false;

                // Keep the existing FitToPagesWide value (do not modify it)
                // Set FitToPagesTall to 1 so each sheet fits vertically on a single page
                ps.FitToPagesTall = 1;

                // Optional: define a print area that covers the used range
                int lastRow = ws.Cells.MaxDataRow;
                int lastCol = ws.Cells.MaxDataColumn;
                ps.PrintArea = $"A1:{CellIndexToName(lastRow, lastCol)}";
            }

            // -------------------------------------------------
            // Save the workbook as a multi‑sheet PDF.
            // The scaling settings applied above are respected
            // for every sheet, giving consistent output.
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Do not force one page per sheet; allow multiple pages
                // while preserving the FitToPagesTall = 1 setting.
                OnePagePerSheet = false,
                // Keep column width handling default (no forced single‑page width)
                AllColumnsInOnePagePerSheet = false
            };

            workbook.Save("MultiSheetOutput.pdf", pdfOptions);
        }

        // Helper to convert zero‑based row/column indexes to Excel cell name (e.g., 0,0 -> A1)
        private static string CellIndexToName(int row, int col)
        {
            string columnName = "";
            int dividend = col + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return $"{columnName}{row + 1}";
        }
    }
}
