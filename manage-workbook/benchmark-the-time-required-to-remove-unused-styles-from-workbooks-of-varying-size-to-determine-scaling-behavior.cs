using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsStyleBenchmark
{
    class Program
    {
        static void Main()
        {
            // Define workbook sizes to test (number of rows and columns)
            int[] rowCounts = { 500, 1000, 2000 };
            int colCount = 10; // fixed number of columns

            // For each size, create a workbook, apply many distinct styles,
            // delete some rows to make styles unused, then benchmark RemoveUnusedStyles.
            foreach (int rows in rowCounts)
            {
                // Create workbook and populate with data and unique styles
                Workbook wb = CreateWorkbookWithUniqueStyles(rows, colCount);

                // Record number of styles before removal
                int styleCountBefore = wb.CountOfStylesInPool;

                // Delete half of the rows to make many styles unused
                int rowsToDelete = rows / 2;
                wb.Worksheets[0].Cells.DeleteRows(rowsToDelete, rows - rowsToDelete);

                // Benchmark RemoveUnusedStyles
                Stopwatch sw = Stopwatch.StartNew();
                wb.RemoveUnusedStyles();
                sw.Stop();

                // Record number of styles after removal
                int styleCountAfter = wb.CountOfStylesInPool;

                // Output results
                Console.WriteLine($"Rows: {rows}, Columns: {colCount}");
                Console.WriteLine($"Styles before removal: {styleCountBefore}");
                Console.WriteLine($"Styles after removal:  {styleCountAfter}");
                Console.WriteLine($"RemoveUnusedStyles elapsed: {sw.ElapsedMilliseconds} ms");
                Console.WriteLine(new string('-', 50));

                // Optionally save the workbook for inspection (uses provided Save method)
                string fileName = $"Benchmark_{rows}x{colCount}.xlsx";
                wb.Save(fileName);
                wb.Dispose();
            }
        }

        // Creates a workbook with the specified number of rows and columns.
        // Each cell receives a distinct style to increase the style pool size.
        static Workbook CreateWorkbookWithUniqueStyles(int rows, int cols)
        {
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Apply a unique style to each cell
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    // Put some sample value
                    cells[r, c].PutValue($"R{r}C{c}");

                    // Create a new style
                    Style style = wb.CreateStyle();

                    // Vary font size and color based on row/column to ensure uniqueness
                    style.Font.Size = 10 + (r % 10);
                    style.Font.Color = System.Drawing.Color.FromArgb(
                        255,
                        (r * 5) % 256,
                        (c * 15) % 256,
                        ((r + c) * 20) % 256);

                    // Apply the style to the cell
                    cells[r, c].SetStyle(style);
                }
            }

            return wb;
        }
    }
}