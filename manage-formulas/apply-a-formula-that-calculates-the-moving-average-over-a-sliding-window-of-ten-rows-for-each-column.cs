using System;
using Aspose.Cells;

namespace MovingAverageExample
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 2. Populate sample data (4 columns, 20 rows)
            int totalRows = 20;
            int totalColumns = 4; // A, B, C, D
            Random rnd = new Random();

            for (int row = 0; row < totalRows; row++)
            {
                for (int col = 0; col < totalColumns; col++)
                {
                    // Fill each cell with a random integer between 1 and 100
                    cells[row, col].PutValue(rnd.Next(1, 101));
                }
            }

            // 3. Apply a 10‑row moving average for each column
            //    The first average will be placed in row 10 (index 9) and will use rows 1‑10.
            //    Subsequent rows will automatically shift the reference when we use a shared formula.
            int windowSize = 10;                     // sliding window length
            int firstAvgRowIndex = windowSize - 1;   // zero‑based index of the first cell that will contain the average

            for (int col = 0; col < totalColumns; col++)
            {
                // Number of rows that will receive the moving‑average formula
                int rowsWithFormula = totalRows - firstAvgRowIndex;
                if (rowsWithFormula <= 0) continue; // not enough data for this column

                // Column letter (e.g., "A", "B", ...)
                string colLetter = CellsHelper.ColumnIndexToName(col);

                // Address of the first cell that will hold the average (e.g., "A10")
                string firstCellAddress = $"{colLetter}{firstAvgRowIndex + 1}";

                // Formula for the first cell: =AVERAGE(A1:A10)
                // Because the formula will be shared, the relative references will shift down automatically.
                string formula = $"=AVERAGE({colLetter}1:{colLetter}{firstAvgRowIndex + 1})";

                // Set the shared formula for the range firstCellAddress : last cell in the column
                cells[firstCellAddress].SetSharedFormula(formula, rowsWithFormula, 1);
            }

            // 4. Calculate all formulas so that the moving averages are evaluated
            workbook.CalculateFormula();

            // 5. Save the workbook
            workbook.Save("MovingAverageResult.xlsx");
        }
    }
}