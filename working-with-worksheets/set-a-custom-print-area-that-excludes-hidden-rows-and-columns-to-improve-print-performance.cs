// Title: C# – Set Custom Print Area Excluding Hidden Rows and Columns with Aspose.Cells
// Description: Demonstrates how to hide rows/columns, detect the first and last visible cells, build an Excel‑style range, assign it to Worksheet.PageSetup.PrintArea, and save the workbook so only visible data is printed.
// Keywords: Aspose.Cells print area C# | exclude hidden rows Aspose.Cells | exclude hidden columns Aspose.Cells | dynamic print range .NET | Worksheet.PageSetup.PrintArea example
// Common Searches: Aspose.Cells set print area without hidden rows | C# get visible range for Excel print area | how to skip hidden columns when printing with Aspose.Cells | dynamic print area based on visible cells .NET
// Developer Intent: Create a print area that contains only the visible rows and columns to reduce page count and improve rendering speed.
// Use Cases: Generate printable reports that omit user‑hidden data. | Export worksheets to PDF while preserving layout of visible cells. | Automatically adjust the print range after interactive hide/show actions.
// AI Prompts: Write C# code using Aspose.Cells to define a print area that automatically skips hidden rows and columns. | Provide a reusable method that returns the Excel range string for the visible portion of a worksheet. | Explain why limiting the print area to visible cells can boost performance in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaDemo
{
    // Demonstrates how to hide rows/columns, detect the first and last visible cells, build an Excel‑style range, assign it to Worksheet.PageSetup.PrintArea, and save the workbook so only visible data is printed.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (10 rows x 5 columns)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Hide some rows and columns to simulate a sparse sheet
            // Hide rows 3 and 7 (zero‑based indices 2 and 6)
            sheet.Cells.HideRow(2);
            sheet.Cells.HideRow(6);
            // Hide columns B and D (zero‑based indices 1 and 3)
            sheet.Cells.HideColumn(1);
            sheet.Cells.HideColumn(3);

            // Determine the first and last visible rows
            int firstVisibleRow = -1, lastVisibleRow = -1;
            for (int r = 0; r < cells.MaxDataRow + 1; r++)
            {
                if (!sheet.Cells.IsRowHidden(r))
                {
                    if (firstVisibleRow == -1) firstVisibleRow = r;
                    lastVisibleRow = r;
                }
            }

            // Determine the first and last visible columns
            int firstVisibleCol = -1, lastVisibleCol = -1;
            for (int c = 0; c < cells.MaxDataColumn + 1; c++)
            {
                if (!sheet.Cells.IsColumnHidden(c))
                {
                    if (firstVisibleCol == -1) firstVisibleCol = c;
                    lastVisibleCol = c;
                }
            }

            // Guard against the case where everything is hidden
            if (firstVisibleRow == -1 || firstVisibleCol == -1)
            {
                Console.WriteLine("All rows or columns are hidden; no print area can be set.");
                return;
            }

            // Convert column indices to Excel column letters
            string startColLetter = CellsHelper.ColumnIndexToName(firstVisibleCol);
            string endColLetter   = CellsHelper.ColumnIndexToName(lastVisibleCol);

            // Build the print area string (Excel uses 1‑based row numbers)
            string printArea = $"{startColLetter}{firstVisibleRow + 1}:{endColLetter}{lastVisibleRow + 1}";

            // Assign the custom print area to the worksheet
            sheet.PageSetup.PrintArea = printArea;

            // Save the workbook (the print area will be respected when printing or exporting)
            workbook.Save("CustomPrintArea.xlsx");

            Console.WriteLine($"Custom print area set to {printArea} and workbook saved.");
        }
    }
}
