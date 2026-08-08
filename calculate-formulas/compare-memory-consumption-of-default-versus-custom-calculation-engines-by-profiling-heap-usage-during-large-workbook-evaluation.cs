// Title: Compare Memory Consumption of Default vs Custom Calculation Engines in Aspose.Cells (C#)
// Description: A C# example that builds two identical large workbooks (5,000 rows × 20 columns) with SUM formulas, measures process memory before and after Workbook.CalculateFormula using the built‑in engine and a no‑op custom AbstractCalculationEngine, reports the heap delta, and saves the results to XLSX files.
// Keywords: Aspose.Cells | memory profiling | calculation engine | custom engine | default engine | C# | .NET | heap usage | formula evaluation | performance benchmark | large workbook | GC | Process.PrivateMemorySize64 | AbstractCalculationEngine
// Common Searches: Aspose.Cells memory usage default engine | custom calculation engine memory benchmark Aspose.Cells | profile heap size during Workbook.CalculateFormula C# | measure private memory increase after formula calculation | compare default and custom calculation engines Aspose.Cells
// Developer Intent: The developer wants to quantify and compare the heap memory impact of the built‑in calculation engine versus a custom AbstractCalculationEngine when processing a large spreadsheet.
// Use Cases: Determine whether a custom calculation engine reduces memory overhead for high‑volume Excel processing. | Benchmark formula evaluation memory consumption to guide performance‑tuning decisions. | Validate that a custom engine does not introduce unacceptable memory growth before production rollout.
// AI Prompts: Create a reusable C# method that captures GC.GetTotalMemory and Process.PrivateMemorySize64 before and after Workbook.CalculateFormula and returns a detailed memory report. | Show how to implement an AbstractCalculationEngine that logs each formula evaluation and records per‑cell memory usage. | Generate a PowerShell script that runs multiple iterations of default and custom engine calculations, aggregates memory deltas, and outputs a comparative chart.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryProfiling
{
    // Custom engine that does not interfere with built‑in functions.
    // It simply overrides Calculate and leaves the calculation to the default engine.
    // A C# example that builds two identical large workbooks (5,000 rows × 20 columns) with SUM formulas, measures process memory before and after Workbook.CalculateFormula using the built‑in engine and a no‑op custom AbstractCalculationEngine, reports the heap delta, and saves the results to XLSX files.
    class NoOpEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // No custom processing – let the default engine handle the function.
        }
    }

    class Program
    {
        static void Main()
        {
            const int rows = 5000;   // large number of rows to stress the engine
            const int cols = 20;     // number of columns with formulas

            // -------------------------------------------------
            // 1. Create a large workbook with formulas
            // -------------------------------------------------
            Workbook wbDefault = new Workbook();
            Worksheet sheetDefault = wbDefault.Worksheets[0];
            Cells cellsDefault = sheetDefault.Cells;

            // Fill cells with numeric values in column A
            for (int r = 0; r < rows; r++)
                cellsDefault[r, 0].PutValue(r + 1);

            // In each subsequent column place a formula that sums the first column up to the current row
            for (int c = 1; c < cols; c++)
            {
                for (int r = 0; r < rows; r++)
                {
                    // Example formula: =SUM($A$1:A{r+1})
                    string formula = $"=SUM($A$1:A{r + 1})";
                    cellsDefault[r, c].Formula = formula;
                }
            }

            // -------------------------------------------------
            // 2. Measure memory usage with the default engine
            // -------------------------------------------------
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long beforeDefault = Process.GetCurrentProcess().PrivateMemorySize64;

            wbDefault.CalculateFormula();   // default calculation

            GC.Collect();
            GC.WaitForPendingFinalizers();
            long afterDefault = Process.GetCurrentProcess().PrivateMemorySize64;

            Console.WriteLine($"Default engine memory increase: {(afterDefault - beforeDefault) / (1024.0 * 1024.0):F2} MB");

            // -------------------------------------------------
            // 3. Create another workbook with the same data
            // -------------------------------------------------
            Workbook wbCustom = new Workbook();
            Worksheet sheetCustom = wbCustom.Worksheets[0];
            Cells cellsCustom = sheetCustom.Cells;

            for (int r = 0; r < rows; r++)
                cellsCustom[r, 0].PutValue(r + 1);

            for (int c = 1; c < cols; c++)
            {
                for (int r = 0; r < rows; r++)
                {
                    string formula = $"=SUM($A$1:A{r + 1})";
                    cellsCustom[r, c].Formula = formula;
                }
            }

            // -------------------------------------------------
            // 4. Measure memory usage with a custom (no‑op) engine
            // -------------------------------------------------
            CalculationOptions options = new CalculationOptions
            {
                CustomEngine = new NoOpEngine()
            };

            GC.Collect();
            GC.WaitForPendingFinalizers();
            long beforeCustom = Process.GetCurrentProcess().PrivateMemorySize64;

            wbCustom.CalculateFormula(options);   // calculation using custom engine

            GC.Collect();
            GC.WaitForPendingFinalizers();
            long afterCustom = Process.GetCurrentProcess().PrivateMemorySize64;

            Console.WriteLine($"Custom engine memory increase: {(afterCustom - beforeCustom) / (1024.0 * 1024.0):F2} MB");

            // -------------------------------------------------
            // 5. Save both workbooks (demonstrates lifecycle usage)
            // -------------------------------------------------
            wbDefault.Save("DefaultEngineResult.xlsx");
            wbCustom.Save("CustomEngineResult.xlsx");

            Console.WriteLine("Profiling completed. Workbooks saved.");
        }
    }
}
