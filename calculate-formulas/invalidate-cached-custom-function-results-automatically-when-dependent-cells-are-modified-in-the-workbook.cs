// Title: Invalidate Cached Results of a Custom Volatile Function in Aspose.Cells for .NET
// Description: Demonstrates how to force recalculation of a user‑defined function by overriding AbstractCalculationEngine.ForceRecalculate, enabling the calculation chain, and using CalculationOptions so that cached results are cleared whenever dependent cells are edited.
// Keywords: Aspose.Cells custom function | cache invalidation | volatile user defined function | ForceRecalculate | AbstractCalculationEngine | EnableCalculationChain | CalculateFormula | C# Excel API | Aspose.Cells .NET example | Excel custom formula engine
// Common Searches: Aspose.Cells invalidate custom function cache | force recalculate user defined function Aspose.Cells | make custom formula volatile in Aspose.Cells .NET | enable calculation chain for dependent cells Aspose.Cells | how to refresh cached results of a custom Excel function
// Developer Intent: Automatically clear the cached value of a custom function whenever any of its input cells are modified.
// Use Cases: Mark a specific user‑defined function as volatile so it updates on each dependent cell change. | Track cell dependencies with EnableCalculationChain and recalculate formulas via CalculationOptions. | Log each invocation of the custom function to verify that the cache is being invalidated after edits.
// AI Prompts: Create a C# AbstractCalculationEngine that forces recalculation of a named custom function whenever its arguments change in Aspose.Cells. | Show how to enable the calculation chain and use CalculationOptions to automatically invalidate cached results of a volatile user‑defined function. | Provide sample code that counts calls to a custom function and demonstrates cache clearing after updating a referenced cell.

using System;
using Aspose.Cells;

namespace CustomFunctionCacheInvalidationDemo
{
    // Custom calculation engine that forces recalculation of the custom function
    // Demonstrates how to force recalculation of a user‑defined function by overriding AbstractCalculationEngine.ForceRecalculate, enabling the calculation chain, and using CalculationOptions so that cached results are cleared whenever dependent cells are edited.
    public class MyCustomEngine : AbstractCalculationEngine
    {
        // Force recalculation for the custom function "MYVOLATILEFUNC"
        public override bool ForceRecalculate(string functionName)
        {
            return string.Equals(functionName, "MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase);
        }

        // Simple implementation: return the value of the first parameter plus a counter
        private int _callCount = 0;
        public override void Calculate(CalculationData data)
        {
            if (string.Equals(data.FunctionName, "MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Increment call counter to show that the function is re‑executed
                _callCount++;

                // Get the first parameter value (expected to be a numeric cell value)
                object param = data.GetParamValue(0);
                double input = Convert.ToDouble(param);

                // Example calculation: input value multiplied by call count
                data.CalculatedValue = input * _callCount;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Enable calculation chain so dependent cells are tracked
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Place an initial value in A1 (the dependent cell)
            cells["A1"].PutValue(5);

            // Set a formula that uses the custom volatile function
            cells["B1"].Formula = "=MYVOLATILEFUNC(A1)";

            // Prepare calculation options with the custom engine
            CalculationOptions calcOptions = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // First calculation – should invoke the custom function once
            workbook.CalculateFormula(calcOptions);
            Console.WriteLine($"After first calc, B1 = {cells["B1"].Value} (expected 5)");

            // Modify the dependent cell A1
            cells["A1"].PutValue(10);

            // Re‑calculate – because ForceRecalculate returns true,
            // the cached result for MYVOLATILEFUNC is invalidated automatically
            workbook.CalculateFormula(calcOptions);
            Console.WriteLine($"After modifying A1, B1 = {cells["B1"].Value} (expected 20)");

            // Change A1 again to demonstrate further invalidation
            cells["A1"].PutValue(2);
            workbook.CalculateFormula(calcOptions);
            Console.WriteLine($"After second modification, B1 = {cells["B1"].Value} (expected 6)");

            // ---------- Save the workbook ----------
            workbook.Save("CustomFunctionCacheInvalidationDemo.xlsx");
        }
    }
}
