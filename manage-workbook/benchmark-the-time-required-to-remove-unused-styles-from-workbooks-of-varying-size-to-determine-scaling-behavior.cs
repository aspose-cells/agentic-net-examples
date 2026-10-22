// Title: How to benchmark the execution time of Aspose.Cells RemoveUnusedStyles on workbooks of different sizes in C#
// Common Searches: C# benchmark RemoveUnusedStyles method for large Excel workbooks using Aspose.Cells | measure performance of Aspose.Cells style cleanup as workbook rows and columns increase | determine scaling of RemoveUnusedStyles with workbook dimensions in .NET | time how long Aspose.Cells takes to delete unused styles in a 2000x2000 sheet | profile Aspose.Cells RemoveUnusedStyles for varying workbook sizes in C#
// Tags: Aspose.Cells RemoveUnusedStyles performance benchmark | C# workbook style cleanup timing | scaling analysis of style removal in large Excel files | measure Aspose.Cells style removal latency .NET | benchmarking workbook size impact on RemoveUnusedStyles | performance testing of Aspose.Cells style optimization

using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsBenchmark
{
    // The sample creates workbooks of several dimensions, applies a pool of random styles to every cell, optionally warms up the workbook, then measures and outputs the elapsed milliseconds for the RemoveUnusedStyles call, allowing developers to see how the cleanup time scales with workbook size.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define workbook sizes to test (rows x columns)
                var sizes = new List<(int rows, int cols)>
                {
                    (100, 100),
                    (500, 500),
                    (1000, 1000),
                    (2000, 2000)
                };

                // Number of distinct styles to apply before cleanup
                const int styleCount = 50;

                foreach (var (rows, cols) in sizes)
                {
                    // Create workbook and populate with data and styles
                    var workbook = CreateWorkbook(rows, cols, styleCount);

                    // Warm up (optional)
                    workbook.CalculateFormula();

                    // Benchmark RemoveUnusedStyles
                    var sw = Stopwatch.StartNew();
                    workbook.RemoveUnusedStyles();
                    sw.Stop();

                    Console.WriteLine(
                        $"Size: {rows}x{cols}, Styles Applied: {styleCount}, RemoveUnusedStyles Time: {sw.ElapsedMilliseconds} ms");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Creates a workbook with the specified number of rows and columns,
        // and applies a set of random styles to cells.
        private static Workbook CreateWorkbook(int rows, int cols, int styleCount)
        {
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Generate a pool of distinct styles
            var styles = new List<Style>();
            var random = new Random();

            for (int i = 0; i < styleCount; i++)
            {
                var s = workbook.CreateStyle();

                // Randomly set some style properties
                s.Font.Color = System.Drawing.Color.FromArgb(
                    random.Next(256), random.Next(256), random.Next(256));
                s.Font.IsBold = random.Next(2) == 0;
                s.ForegroundColor = System.Drawing.Color.FromArgb(
                    random.Next(256), random.Next(256), random.Next(256));
                s.Pattern = random.Next(2) == 0 ? BackgroundType.Solid : BackgroundType.None;

                // The 'Custom' property is read‑only or of a different type; omit setting it.
                styles.Add(s);
            }

            // Populate cells with data and assign random styles
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    var cell = sheet.Cells[r, c];
                    cell.PutValue($"R{r}C{c}");

                    // Assign a random style from the pool
                    int styleIndex = random.Next(styleCount);
                    cell.SetStyle(styles[styleIndex]);
                }
            }

            return workbook;
        }
    }
}
