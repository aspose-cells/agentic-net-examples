// Title: Memory Profiling: Default vs Custom Calculation Engine in Aspose.Cells for .NET
// Description: Generates a 5,000‑row workbook with formulas, records heap size before and after CalculateFormula using the built‑in engine, repeats the measurement with a PassThroughEngine custom engine, prints the memory delta, and saves the workbook.
// Keywords: Aspose.Cells | C# memory profiling | .NET heap usage | calculation engine performance | custom calculation engine | default engine memory impact | large workbook formula evaluation | GC.GetTotalMemory | performance measurement | formula recalculation
// Common Searches: Aspose.Cells memory profiling example | compare default and custom calculation engine memory usage | heap usage during CalculateFormula in C# | measure memory impact of custom engine Aspose.Cells | performance test for large workbook calculations .NET
// Developer Intent: The developer wants to quantify the heap memory difference between Aspose.Cells' built‑in calculation engine and a user‑defined engine when processing a large set of formulas.
// Use Cases: Determine if a custom calculation engine introduces additional memory overhead in high‑volume spreadsheets. | Identify memory bottlenecks during bulk formula recalculation. | Validate that a lightweight pass‑through engine consumes comparable memory to the default engine.
// AI Prompts: Create C# code that logs detailed memory statistics (pre‑ and post‑GC) for each CalculateFormula call in Aspose.Cells. | Show how to extend AbstractCalculationEngine to handle specific functions while still measuring memory consumption. | Suggest best practices to minimize heap allocation when evaluating formulas in large Aspose.Cells workbooks.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryProfiling
{
    // Simple custom engine that does not process any built‑in functions.
    // The default engine will handle all calculations.
    // Generates a 5,000‑row workbook with formulas, records heap size before and after CalculateFormula using the built‑in engine, repeats the measurement with a PassThroughEngine custom engine, prints the memory delta, and saves the workbook.
    public class PassThroughEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // No custom processing; let the default engine handle the function.
        }
    }

    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a large workbook with many formulas
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            const int rows = 5000;   // adjust for desired size
            const int cols = 10;

            // Populate column A with numeric values
            for (int r = 0; r < rows; r++)
            {
                cells[r, 0].PutValue(r + 1);
            }

            // Populate other columns with formulas that depend on column A
            for (int c = 1; c < cols; c++)
            {
                for (int r = 0; r < rows; r++)
                {
                    // Example formula: =A1*2 + COLUMN_INDEX
                    string formula = $"=A{r + 1}*2+{c}";
                    cells[r, c].Formula = formula;
                }
            }

            // ------------------------------------------------------------
            // 2. Measure memory usage with the default calculation engine
            // ------------------------------------------------------------
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memBeforeDefault = GC.GetTotalMemory(true);

            // Calculate all formulas using the built‑in engine
            workbook.CalculateFormula();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memAfterDefault = GC.GetTotalMemory(true);
            long usedDefault = memAfterDefault - memBeforeDefault;

            // ------------------------------------------------------------
            // 3. Measure memory usage with a custom calculation engine
            // ------------------------------------------------------------
            CalculationOptions customOptions = new CalculationOptions
            {
                CustomEngine = new PassThroughEngine()
            };

            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memBeforeCustom = GC.GetTotalMemory(true);

            // Re‑calculate using the custom engine
            workbook.CalculateFormula(customOptions);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memAfterCustom = GC.GetTotalMemory(true);
            long usedCustom = memAfterCustom - memBeforeCustom;

            // ------------------------------------------------------------
            // 4. Output the profiling results
            // ------------------------------------------------------------
            Console.WriteLine($"Memory used by default engine : {usedDefault:N0} bytes");
            Console.WriteLine($"Memory used by custom engine  : {usedCustom:N0} bytes");
            Console.WriteLine($"Difference (custom - default) : {usedCustom - usedDefault:N0} bytes");

            // ------------------------------------------------------------
            // 5. Save the workbook (demonstrates the required save lifecycle)
            // ------------------------------------------------------------
            workbook.Save("MemoryProfilingResult.xlsx");
        }
    }
}
