// Title: Aspose.Cells for .NET – Set Print Area Excluding Hidden Rows & Columns (C#)
// Description: C# example that creates a workbook, hides selected rows and columns, determines the first and last visible cells, builds a range address, assigns it to Worksheet.PageSetup.PrintArea, and saves the file. The dynamic print area improves printing speed by omitting hidden data.
// Keywords: Aspose.Cells print area C# | exclude hidden rows Aspose.Cells | custom print range .NET | Worksheet.PageSetup.PrintArea | dynamic print area visible cells | C# Excel hide rows columns | Aspose.Cells performance printing
// Common Searches: how to set print area in Aspose.Cells ignoring hidden rows | Aspose.Cells C# set print area based on visible cells | exclude hidden columns from print area Aspose.Cells | dynamic print range for worksheet Aspose.Cells | C# Aspose.Cells print area after hiding rows
// Developer Intent: Create a print area that contains only the visible rows and columns to reduce file size and speed up printing.
// Use Cases: Generate printable reports that automatically skip hidden sections. | Prepare workbooks for batch printing where hidden data should not appear. | Design templates that adapt their print range after users hide rows or columns.
// AI Prompts: Modify the code to also ignore rows and columns hidden by auto‑filter when setting the print area. | Show how to apply the same visible‑range logic to every worksheet in a multi‑sheet workbook. | Replace the manual loops with Worksheet.Cells.MaxDisplayRange (or a similar API) to compute the visible range for PrintArea.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaDemo
{
    // C# example that creates a workbook, hides selected rows and columns, determines the first and last visible cells, builds a range address, assigns it to Worksheet.PageSetup.PrintArea, and saves the file. The dynamic print area improves printing speed by omitting hidden data.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (A1 to E10)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Hide some rows and columns to simulate a sparse sheet
            sheet.Cells.HideRow(2);   // hide row 3 (zero‑based index)
            sheet.Cells.HideRow(5);   // hide row 6
            sheet.Cells.HideColumn(1); // hide column B
            sheet.Cells.HideColumn(3); // hide column D

            // Determine the first and last visible rows
            int firstVisibleRow = -1, lastVisibleRow = -1;
            for (int r = 0; r < sheet.Cells.MaxDataRow + 1; r++)
            {
                if (!sheet.Cells.IsRowHidden(r))
                {
                    if (firstVisibleRow == -1) firstVisibleRow = r;
                    lastVisibleRow = r;
                }
            }

            // Determine the first and last visible columns
            int firstVisibleCol = -1, lastVisibleCol = -1;
            for (int c = 0; c < sheet.Cells.MaxDataColumn + 1; c++)
            {
                if (!sheet.Cells.IsColumnHidden(c))
                {
                    if (firstVisibleCol == -1) firstVisibleCol = c;
                    lastVisibleCol = c;
                }
            }

            // Guard against completely hidden sheet
            if (firstVisibleRow == -1 || firstVisibleCol == -1)
            {
                Console.WriteLine("All rows or columns are hidden. No print area will be set.");
                return;
            }

            // Build the address string for the print area (e.g., "A1:E10")
            string startCell = GetCellName(firstVisibleRow, firstVisibleCol);
            string endCell   = GetCellName(lastVisibleRow, lastVisibleCol);
            string printArea = $"{startCell}:{endCell}";

            // Set the custom print area (property rule)
            sheet.PageSetup.PrintArea = printArea;

            // Save the workbook (save rule)
            workbook.Save("CustomPrintArea.xlsx");

            Console.WriteLine($"Print area set to {printArea} and workbook saved.");
        }

        // Helper: converts zero‑based row/column indexes to Excel cell name (e.g., 0,0 -> "A1")
        private static string GetCellName(int rowIndex, int columnIndex)
        {
            // Convert column index to letters
            string columnName = "";
            int dividend = columnIndex + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            // Row index is zero‑based; Excel rows start at 1
            int rowNumber = rowIndex + 1;
            return $"{columnName}{rowNumber}";
        }
    }
}
