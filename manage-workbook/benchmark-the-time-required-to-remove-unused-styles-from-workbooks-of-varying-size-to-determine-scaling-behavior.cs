// Title: Benchmarking Aspose.Cells RemoveUnusedStyles method on C# workbooks of varying row counts
// Common Searches: C# benchmark RemoveUnusedStyles execution time for a 10,000 row Excel workbook using Aspose.Cells | how to measure Aspose.Cells style cleanup performance with different workbook sizes | profile RemoveUnusedStyles scaling behavior in large Excel files with Aspose.Cells C# | time taken to delete unused styles in a 20,000 row workbook using Aspose.Cells | Aspose.Cells RemoveUnusedStyles speed test for varying numbers of rows in C#
// Tags: remove unused styles performance Aspose.Cells | benchmark RemoveUnusedStyles C# | style cleanup scaling Aspose.Cells | large workbook style optimization Aspose.Cells | measure RemoveUnusedStyles execution time | Aspose.Cells workbook performance testing

using System;
using System.Diagnostics;
using System.Drawing;
using Aspose.Cells;

// The sample creates workbooks ranging from 1,000 to 20,000 rows, applies random styles to each cell, adds 500 unused styles, and measures the elapsed milliseconds of workbook.RemoveUnusedStyles() for each size, providing a clear scaling benchmark of the style‑cleanup operation.
class RemoveUnusedStylesBenchmark
{
    static void Main()
    {
        try
        {
            // Define workbook sizes (number of rows) to test
            int[] rowCounts = new int[] { 1000, 5000, 10000, 20000 };

            // Number of columns to fill (fixed)
            int columnCount = 10;

            // Random generator used for style variations
            Random rnd = new Random();

            foreach (int rows in rowCounts)
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the worksheet with data and apply random styles
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < columnCount; c++)
                    {
                        // Write a simple value
                        sheet.Cells[r, c].PutValue($"R{r}C{c}");

                        // Create a new style and randomize its appearance
                        Style style = workbook.CreateStyle();
                        style.Font.Color = Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));
                        style.Font.Size = rnd.Next(8, 15);               // Font.Size expects an int
                        style.ForegroundColor = Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));
                        style.Pattern = BackgroundType.Solid;
                        style.Font.IsBold = rnd.Next(2) == 0;            // Use Font.IsBold instead of Style.IsBold

                        // Assign the style to the cell
                        sheet.Cells[r, c].SetStyle(style);
                    }
                }

                // Create additional styles that are never applied (to simulate unused styles)
                for (int i = 0; i < 500; i++)
                {
                    Style unusedStyle = workbook.CreateStyle();
                    unusedStyle.Font.Color = Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));
                    unusedStyle.Font.Size = rnd.Next(8, 15);
                    // Do NOT assign this style to any cell
                }

                // Measure the time taken to remove unused styles
                Stopwatch sw = Stopwatch.StartNew();
                workbook.RemoveUnusedStyles();
                sw.Stop();

                // Output the benchmark result
                Console.WriteLine($"Rows: {rows}, Columns: {columnCount}, Time to remove unused styles: {sw.ElapsedMilliseconds} ms");

                // Optional: Save the workbook to verify the operation (can be omitted for pure benchmarking)
                // string fileName = $"Benchmark_{rows}Rows.xlsx";
                // workbook.Save(fileName);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
