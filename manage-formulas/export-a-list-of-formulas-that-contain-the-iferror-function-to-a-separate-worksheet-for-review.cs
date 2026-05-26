using System;
using Aspose.Cells;

namespace AsposeCellsIFERRORExport
{
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with actual path)
            Workbook workbook = new Workbook("source.xlsx");

            // Prepare a new worksheet to hold the IFERROR formulas
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet reviewSheet = workbook.Worksheets[newSheetIndex];
            Cells reviewCells = reviewSheet.Cells;

            // Write header row
            reviewCells[0, 0].PutValue("Worksheet");
            reviewCells[0, 1].PutValue("Cell Address");
            reviewCells[0, 2].PutValue("Formula");

            int outputRow = 1; // start after header

            // Iterate through all worksheets in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Cells cells = ws.Cells;

                // Determine the used range
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Check if the cell has a formula
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            // Look for the IFERROR function (case‑insensitive)
                            if (cell.Formula.IndexOf("IFERROR", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Record the worksheet name, cell address and the formula
                                reviewCells[outputRow, 0].PutValue(ws.Name);
                                reviewCells[outputRow, 1].PutValue(cell.Name);
                                reviewCells[outputRow, 2].PutValue(cell.Formula);
                                outputRow++;
                            }
                        }
                    }
                }
            }

            // Save the workbook with the new review sheet (replace with desired output path)
            workbook.Save("output_with_iferror_review.xlsx");
        }
    }
}