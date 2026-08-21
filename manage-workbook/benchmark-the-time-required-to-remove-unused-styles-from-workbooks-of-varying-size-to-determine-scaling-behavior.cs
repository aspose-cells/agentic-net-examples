// Title: Benchmark RemoveUnusedStyles Performance in Aspose.Cells C# Across Workbook Sizes
// Description: C# sample that creates workbooks with 500, 2,000 and 5,000 rows, assigns a distinct style to each row, deletes half of the rows to leave many styles unused, then measures the elapsed time and style count before and after calling Workbook.RemoveUnusedStyles().
// Keywords: Aspose.Cells RemoveUnusedStyles benchmark | style cleanup performance .NET | measure Aspose.Cells style removal time | Workbook.RemoveUnusedStyles scaling | C# Aspose.Cells performance testing
// Common Searches: How fast is RemoveUnusedStyles on large Aspose.Cells workbooks | Aspose.Cells style removal benchmark C# | Performance of Workbook.RemoveUnusedStyles with many rows | Timing unused style cleanup in Aspose.Cells | Scale of RemoveUnusedStyles execution time
// Developer Intent: Assess how the execution time of Workbook.RemoveUnusedStyles varies with the number of rows and styles in a workbook.
// Use Cases: Validate that style cleanup meets latency requirements for automated report generation. | Compare removal times for small, medium, and large spreadsheets before production deployment. | Identify whether additional optimization is needed when handling workbooks with thousands of unique styles.
// AI Prompts: Generate a reusable C# helper that runs RemoveUnusedStyles on a Workbook and returns the elapsed milliseconds. | Extend the benchmark to record peak memory usage during style removal. | Explain which workbook characteristics (row count, style count, formula complexity) most affect RemoveUnusedStyles performance and how to interpret the results.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsStyleBenchmark
{
    // C# sample that creates workbooks with 500, 2,000 and 5,000 rows, assigns a distinct style to each row, deletes half of the rows to leave many styles unused, then measures the elapsed time and style count before and after calling Workbook.RemoveUnusedStyles().
    class Program
    {
        static void Main(string[] args)
        {
            // Define different workbook sizes (number of rows to populate)
            int[] rowCounts = new int[] { 500, 2000, 5000 };

            foreach (int rowCount in rowCounts)
            {
                // Create a new workbook
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate cells with unique styles
                for (int i = 0; i < rowCount; i++)
                {
                    // Put a simple value
                    cells[i, 0].PutValue($"Row {i + 1}");

                    // Create a distinct style for each row
                    Style style = wb.CreateStyle();
                    style.Font.Size = 10 + (i % 20);               // Vary font size
                    style.Font.IsBold = (i % 2 == 0);              // Alternate bold
                    style.Font.Color = System.Drawing.Color.FromArgb(255, (i * 5) % 256, (i * 3) % 256);
                    cells[i, 0].SetStyle(style);
                }

                // Record style count before any removal
                int styleCountBefore = wb.CountOfStylesInPool;

                // Delete half of the rows to make many styles unused
                int rowsToDelete = rowCount / 2;
                sheet.Cells.DeleteRows(rowsToDelete, rowsToDelete);

                // Benchmark RemoveUnusedStyles
                Stopwatch sw = Stopwatch.StartNew();
                wb.RemoveUnusedStyles();
                sw.Stop();

                // Record style count after removal
                int styleCountAfter = wb.CountOfStylesInPool;

                // Output results
                Console.WriteLine($"Rows: {rowCount}");
                Console.WriteLine($"  Styles before removal: {styleCountBefore}");
                Console.WriteLine($"  Styles after removal:  {styleCountAfter}");
                Console.WriteLine($"  Time to remove unused styles: {sw.ElapsedMilliseconds} ms");
                Console.WriteLine();

                // Optional: save the workbook for manual inspection (not required for benchmark)
                string fileName = $"Benchmark_{rowCount}_Rows.xlsx";
                wb.Save(fileName);
                wb.Dispose();
            }
        }
    }
}
