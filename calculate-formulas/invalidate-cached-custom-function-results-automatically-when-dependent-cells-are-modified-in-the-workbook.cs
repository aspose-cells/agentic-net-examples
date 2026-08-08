// Title: Invalidate Cached Results of a Custom Volatile Function in Aspose.Cells for .NET
// Description: Demonstrates how to force recalculation of a custom volatile function in Aspose.Cells by enabling the calculation chain, implementing a custom calculation engine with ForceRecalculate, detecting dependent cells, and re‑calculating the workbook after input changes. The example shows cache invalidation, call‑counter verification, and saving the updated file.
// Keywords: Aspose.Cells custom function cache invalidation | volatile custom function Aspose.Cells | ForceRecalculate Aspose.Cells | EnableCalculationChain Aspose.Cells | GetDependentsInCalculation C# | Aspose.Cells custom calculation engine | C# Excel custom function recalc | Aspose.Cells workbook.CalculateFormula
// Common Searches: how to invalidate custom function cache Aspose.Cells | force recalculate custom function in .NET Excel library | enable calculation chain to track dependencies Aspose.Cells | retrieve dependent cells after change Aspose.Cells | volatile custom function example Aspose.Cells C#
// Developer Intent: Make a custom Excel function recompute automatically whenever any of its referenced cells are edited, ensuring the cached result is cleared.
// Use Cases: Create a volatile custom function that adds a call counter to prove re‑execution after dependent cells change. | Enable the calculation chain and use GetDependentsInCalculation to list cells that depend on a modified input. | Recalculate the workbook with a custom engine, verify updated results, and save the workbook.
// AI Prompts: Show C# code for a custom calculation engine that forces recalculation of a specific function in Aspose.Cells. | Explain how to enable the calculation chain and retrieve dependent cells for a given cell using Aspose.Cells. | Provide a step‑by‑step example that demonstrates cache invalidation for a volatile custom function when an input cell value is modified.

using System;
using System.Collections;
using Aspose.Cells;

namespace CustomFunctionCacheInvalidationDemo
{
    // Custom calculation engine that forces recalculation of the custom function
    // Demonstrates how to force recalculation of a custom volatile function in Aspose.Cells by enabling the calculation chain, implementing a custom calculation engine with ForceRecalculate, detecting dependent cells, and re‑calculating the workbook after input changes. The example shows cache invalidation, call‑counter verification, and saving the updated file.
    public class MyCustomEngine : AbstractCalculationEngine
    {
        // Force recalculation for the custom function "MYVOLATILEFUNC"
        public override bool ForceRecalculate(string functionName)
        {
            return string.Equals(functionName, "MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase);
        }

        // Simple implementation: sum of all parameters plus a running counter
        private int _callCount = 0;
        public override void Calculate(CalculationData data)
        {
            if (string.Equals(data.FunctionName, "MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);
                    // Parameters may be simple values or ReferredArea objects
                    if (param is double d)
                    {
                        sum += d;
                    }
                    else if (param is int iVal)
                    {
                        sum += iVal;
                    }
                    else if (param is ReferredArea area)
                    {
                        // Take the first cell of the area for simplicity
                        sum += Convert.ToDouble(area.GetValue(0, 0));
                    }
                }
                _callCount++; // increase call counter to prove re‑execution
                data.CalculatedValue = sum + _callCount;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create ----------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Enable calculation chain so that dependent cells can be discovered
                workbook.Settings.FormulaSettings.EnableCalculationChain = true;

                // Input cells that the custom function will depend on
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);

                // Cell B1 uses the custom volatile function
                cells["B1"].Formula = "=MYVOLATILEFUNC(A1, A2)";

                // ---------- Set up custom engine ----------
                MyCustomEngine customEngine = new MyCustomEngine();
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CustomEngine = customEngine
                };

                // ---------- First calculation ----------
                workbook.CalculateFormula(calcOptions);
                Console.WriteLine($"Initial B1 value: {cells["B1"].Value}");

                // ---------- Modify a dependent cell ----------
                cells["A1"].PutValue(30); // change A1, which B1 depends on

                // Optionally, retrieve dependents of A1 (demonstrates dependency chain)
                IEnumerator dependents = cells.GetDependentsInCalculation(0, 0, true);
                Console.WriteLine("Cells dependent on A1 after change:");
                while (dependents.MoveNext())
                {
                    if (dependents.Current is Cell depCell)
                    {
                        Console.WriteLine($"- {depCell.Name}");
                    }
                }

                // ---------- Re‑calculate ----------
                workbook.CalculateFormula(calcOptions);
                Console.WriteLine($"After A1 change, B1 value: {cells["B1"].Value}");

                // ---------- Save ----------
                string outputPath = "CustomFunctionCacheInvalidationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
