using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (10 rows x 8 columns)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 8; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Hide some rows and columns to simulate a sparse sheet
            worksheet.Cells.HideRow(2);   // hide row 3 (zero‑based index)
            worksheet.Cells.HideRow(5);   // hide row 6
            worksheet.Cells.HideColumn(1); // hide column B
            worksheet.Cells.HideColumn(4); // hide column E

            // Determine the first and last visible rows
            int firstVisibleRow = -1, lastVisibleRow = -1;
            for (int r = 0; r <= worksheet.Cells.MaxDataRow; r++)
            {
                if (!worksheet.Cells.IsRowHidden(r))
                {
                    if (firstVisibleRow == -1) firstVisibleRow = r;
                    lastVisibleRow = r;
                }
            }

            // Determine the first and last visible columns
            int firstVisibleCol = -1, lastVisibleCol = -1;
            for (int c = 0; c <= worksheet.Cells.MaxDataColumn; c++)
            {
                if (!worksheet.Cells.IsColumnHidden(c))
                {
                    if (firstVisibleCol == -1) firstVisibleCol = c;
                    lastVisibleCol = c;
                }
            }

            // Guard against a completely hidden sheet
            if (firstVisibleRow == -1 || firstVisibleCol == -1)
            {
                Console.WriteLine("All rows or columns are hidden. No print area will be set.");
            }
            else
            {
                // Build the address string for the visible area, e.g., "A1:G8"
                string startCell = CellsHelper.CellIndexToName(firstVisibleRow, firstVisibleCol);
                string endCell   = CellsHelper.CellIndexToName(lastVisibleRow, lastVisibleCol);
                string printArea = $"{startCell}:{endCell}";

                // Set the custom print area
                worksheet.PageSetup.PrintArea = printArea;

                Console.WriteLine($"Custom print area set to: {printArea}");
            }

            // Save the workbook (PDF will respect the PrintArea)
            workbook.Save("CustomPrintArea.pdf", SaveFormat.Pdf);
            Console.WriteLine("Workbook saved as CustomPrintArea.pdf");
        }
    }
}