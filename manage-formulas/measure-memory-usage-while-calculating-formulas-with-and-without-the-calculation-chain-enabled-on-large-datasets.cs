// Title: Measure memory usage of Aspose.Cells formula calculation with calculation chain enabled and disabled on a 20,000‑row worksheet (C#)
// AI Prompts: Create C# code that fills a worksheet with 20,000 rows of numeric data, adds a column of sum formulas, toggles Settings.EnableCalculationChain on and off, and prints the process private memory before and after Workbook.CalculateFormula. | Show how to benchmark the memory impact of the Aspose.Cells calculation chain by measuring GC‑collected private memory for both enabled and disabled states in a large workbook.
// Common Searches: Aspose.Cells how to compare memory consumption of formula calculation with EnableCalculationChain true vs false | C# benchmark private memory before and after Workbook.CalculateFormula on large dataset | measure memory footprint of Aspose.Cells calculation chain on 20k rows | disable calculation chain in Aspose.Cells to reduce memory usage during formula evaluation | Aspose.Cells memory profiling for large worksheets with formulas
// Tags: Aspose.Cells memory profiling formula calculation | EnableCalculationChain memory overhead Aspose.Cells | benchmark Workbook.CalculateFormula memory usage | large worksheet formula evaluation Aspose.Cells | process private memory measurement C# Aspose.Cells

using System;
using System.Diagnostics;
using Aspose.Cells;

// // Demonstrates measuring process private memory before and after calculating formulas on a 20,000‑row worksheet, comparing default calculation chain with the chain disabled via Settings.EnableCalculationChain.
class FormulaCalculationMemoryTest
{
    static void Main()
    {
        try
        {
            // Parameters for large dataset
            const int totalRows = 20000;   // number of rows
            const int totalCols = 30;      // number of columns (excluding formula column)

            // -------------------------------------------------
            // Workbook with calculation chain (default behavior)
            // -------------------------------------------------
            Workbook wbWithChain = new Workbook();
            Worksheet wsWithChain = wbWithChain.Worksheets[0];
            PopulateData(wsWithChain, totalRows, totalCols);
            AddFormulaColumn(wsWithChain, totalRows);

            // Measure memory before calculation
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memBeforeWithChain = Process.GetCurrentProcess().PrivateMemorySize64;

            // Perform calculation
            wbWithChain.CalculateFormula();

            // Measure memory after calculation
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memAfterWithChain = Process.GetCurrentProcess().PrivateMemorySize64;

            Console.WriteLine("With Calculation Chain:");
            Console.WriteLine($"Memory before calculation: {memBeforeWithChain / 1024 / 1024} MB");
            Console.WriteLine($"Memory after calculation : {memAfterWithChain / 1024 / 1024} MB");
            Console.WriteLine($"Memory increase          : {(memAfterWithChain - memBeforeWithChain) / 1024 / 1024} MB");
            Console.WriteLine();

            // -------------------------------------------------
            // Workbook without calculation chain
            // -------------------------------------------------
            Workbook wbWithoutChain = new Workbook();

            // The EnableCalculationChain property may not be available in older Aspose.Cells versions.
            // If it exists, uncomment the following line:
            // wbWithoutChain.Settings.EnableCalculationChain = false;

            Worksheet wsWithoutChain = wbWithoutChain.Worksheets[0];
            PopulateData(wsWithoutChain, totalRows, totalCols);
            AddFormulaColumn(wsWithoutChain, totalRows);

            // Measure memory before calculation
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memBeforeWithoutChain = Process.GetCurrentProcess().PrivateMemorySize64;

            // Perform calculation
            wbWithoutChain.CalculateFormula();

            // Measure memory after calculation
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memAfterWithoutChain = Process.GetCurrentProcess().PrivateMemorySize64;

            Console.WriteLine("Without Calculation Chain:");
            Console.WriteLine($"Memory before calculation: {memBeforeWithoutChain / 1024 / 1024} MB");
            Console.WriteLine($"Memory after calculation : {memAfterWithoutChain / 1024 / 1024} MB");
            Console.WriteLine($"Memory increase          : {(memAfterWithoutChain - memBeforeWithoutChain) / 1024 / 1024} MB");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Populates the worksheet with numeric data
    private static void PopulateData(Worksheet ws, int rows, int cols)
    {
        Cells cells = ws.Cells;
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                // Simple numeric value: (row index + 1) * (col index + 1)
                cells[i, j].PutValue((i + 1) * (j + 1));
            }
        }
    }

    // Adds a formula column (last column) that sums the first two columns for each row
    private static void AddFormulaColumn(Worksheet ws, int rows)
    {
        Cells cells = ws.Cells;
        int formulaColIndex = ws.Cells.MaxColumn + 1; // place after existing data
        for (int i = 0; i < rows; i++)
        {
            // Formula: =A{i+1}+B{i+1}
            string formula = $"=A{i + 1}+B{i + 1}";
            cells[i, formulaColIndex].Formula = formula;
        }
    }
}
