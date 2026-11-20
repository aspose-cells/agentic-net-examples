// Title: Profile memory usage of Aspose.Cells default calculation engine while evaluating a large workbook in C#
// Common Searches: how to measure memory usage of Aspose.Cells CalculateFormula in C# | C# benchmark Aspose.Cells formula calculation heap consumption | profile heap before and after wb.CalculateFormula Aspose.Cells | track garbage collection impact on large workbook evaluation with Aspose.Cells | compare performance of default versus custom calculation engine in Aspose.Cells
// Tags: Aspose.Cells CalculateFormula memory profiling | C# heap measurement Aspose.Cells | large workbook formula evaluation performance | Garbage collection before Aspose.Cells calculation | default calculation engine memory impact

using System;
using System.Diagnostics;
using Aspose.Cells;

// The sample creates a 5,000‑row by 50‑column workbook, fills cells with numeric values, adds a SUM formula per row, forces garbage collection, records total memory before and after calling wb.CalculateFormula (the default engine), and outputs the elapsed time and memory increase in bytes.
class Program
{
    static void Main()
    {
        try
        {
            // Create a large workbook with data and formulas
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            int rows = 5000;
            int cols = 50;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    ws.Cells[i, j].PutValue(i + j);
                    // Add a SUM formula in the last column of each row
                    if (j == cols - 1)
                    {
                        string startCol = GetColumnName(0);
                        string endCol = GetColumnName(cols - 2);
                        ws.Cells[i, j].Formula = $"=SUM({startCol}{i + 1}:{endCol}{i + 1})";
                    }
                }
            }

            // Profile default calculation engine
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long beforeDefault = GC.GetTotalMemory(true);
            Stopwatch sw = Stopwatch.StartNew();
            wb.CalculateFormula(); // default engine evaluates all formulas
            sw.Stop();
            long afterDefault = GC.GetTotalMemory(true);
            Console.WriteLine($"Default engine: Time = {sw.ElapsedMilliseconds} ms, Memory increase = {afterDefault - beforeDefault} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper to convert zero‑based column index to Excel column name (A, B, …, AA, AB, …)
    static string GetColumnName(int index)
    {
        const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        string name = "";
        while (index >= 0)
        {
            name = letters[index % 26] + name;
            index = index / 26 - 1;
        }
        return name;
    }
}
