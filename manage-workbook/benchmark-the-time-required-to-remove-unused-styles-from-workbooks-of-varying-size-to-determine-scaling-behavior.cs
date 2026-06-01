using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class StyleRemovalBenchmark
{
    static void Main()
    {
        // Define different workbook sizes to test.
        var testCases = new (int rows, int cols, int styleCount)[]
        {
            (100, 10, 100),   // Small workbook
            (500, 20, 500),   // Medium workbook
            (1000, 30, 1000)  // Large workbook
        };

        foreach (var (rows, cols, styleCount) in testCases)
        {
            // Create a workbook populated with data and a pool of distinct styles.
            Workbook wb = CreateWorkbook(rows, cols, styleCount);

            // Delete half of the rows to make some styles unused.
            wb.Worksheets[0].Cells.DeleteRows(rows / 2, rows / 2);

            // Record the number of styles before removal.
            int beforeCount = wb.CountOfStylesInPool;
            Console.WriteLine($"Workbook {rows}x{cols}, styles before removal: {beforeCount}");

            // Benchmark the RemoveUnusedStyles method.
            Stopwatch sw = Stopwatch.StartNew();
            wb.RemoveUnusedStyles();
            sw.Stop();

            // Record the number of styles after removal.
            int afterCount = wb.CountOfStylesInPool;
            Console.WriteLine($"After removal: {afterCount} styles, elapsed time: {sw.ElapsedMilliseconds} ms");

            // Save the workbook to a memory stream (satisfies the save rule).
            using (MemoryStream ms = new MemoryStream())
            {
                wb.Save(ms, SaveFormat.Xlsx);
            }

            Console.WriteLine();
        }
    }

    // Creates a workbook with the specified number of rows, columns, and distinct styles.
    static Workbook CreateWorkbook(int rows, int cols, int styleCount)
    {
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        Cells cells = sheet.Cells;

        // Prepare a pool of distinct styles.
        Style[] stylePool = new Style[styleCount];
        Random rnd = new Random(0);
        for (int i = 0; i < styleCount; i++)
        {
            Style style = wb.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 10 + (i % 10);
            style.Font.IsBold = (i % 2 == 0);
            style.Font.Color = Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));
            stylePool[i] = style;
        }

        // Fill cells with data and assign a style from the pool.
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                Cell cell = cells[r, c];
                cell.PutValue($"R{r}C{c}");
                // Cycle through the style pool.
                cell.SetStyle(stylePool[(r * cols + c) % styleCount]);
            }
        }

        return wb;
    }
}