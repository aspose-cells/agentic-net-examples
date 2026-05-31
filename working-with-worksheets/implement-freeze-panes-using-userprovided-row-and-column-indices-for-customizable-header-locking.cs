using System;
using Aspose.Cells;

namespace FreezePaneDemo
{
    public class Program
    {
        // Entry point
        public static void Main()
        {
            // Example user‑provided indices (zero‑based)
            int headerRowIndex = 2;    // Freeze first 3 rows (0,1,2)
            int headerColumnIndex = 1; // Freeze first 2 columns (0,1)

            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data to visualize the effect
            for (int r = 0; r < 20; r++)
            {
                for (int c = 0; c < 10; c++)
                {
                    sheet.Cells[r, c].PutValue($"R{r + 1}C{c + 1}");
                }
            }

            // Freeze panes using the user‑provided indices.
            // The last two parameters define how many rows/columns are frozen;
            // they must not exceed the row/column indices, so we use the same values.
            sheet.FreezePanes(headerRowIndex, headerColumnIndex, headerRowIndex + 1, headerColumnIndex + 1);

            // Optional: verify the freeze settings
            if (sheet.GetFreezedPanes(out int row, out int col, out int frozenRows, out int frozenCols))
            {
                Console.WriteLine($"Freeze position - Row: {row}, Column: {col}");
                Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenCols}");
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("FreezePaneResult.xlsx");
        }
    }
}