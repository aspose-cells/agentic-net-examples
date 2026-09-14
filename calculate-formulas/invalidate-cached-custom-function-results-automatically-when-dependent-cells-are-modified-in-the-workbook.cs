// Title: Automatically invalidate cached results of a custom volatile function when dependent cells change using Aspose.Cells for .NET
// AI Prompts: Create a subclass of AbstractCalculationEngine that forces recalculation for a specific user‑defined function and register it via CalculationOptions in Aspose.Cells. | Enable the calculation chain in a workbook, assign a formula that calls a volatile custom function, then modify a source cell and trigger wb.CalculateFormula to verify the cache is cleared. | Write C# code to sum numeric arguments in a custom function, ensure ForceRecalculate returns true for that function, and save the workbook after the updated calculation.
// Common Searches: Aspose.Cells how to force recalculation of a user defined function after cell update | C# invalidate cached custom function results in Excel workbook using Aspose.Cells | Enable calculation chain and custom calculation engine for volatile functions in Aspose.Cells .NET
// Tags: custom calculation engine Aspose.Cells | force recalculate volatile function .NET | enable calculation chain workbook | invalidate cached formula result C# | user defined function cache clearing Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, enables the calculation chain, defines a volatile custom function (MYVOLATILEFUNC), and implements a custom AbstractCalculationEngine that always forces recalculation for that function. After changing a dependent cell, wb.CalculateFormula with the custom engine clears the cached result and recomputes the value, demonstrating automatic cache invalidation. The workbook is then saved.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Enable calculation chain so dependent cells are tracked
        wb.Settings.FormulaSettings.EnableCalculationChain = true;

        // Populate source cells that the custom function will depend on
        ws.Cells["A1"].PutValue(5);
        ws.Cells["A2"].PutValue(10);

        // Set a formula that uses a custom volatile function
        ws.Cells["B1"].Formula = "=MYVOLATILEFUNC(A1,A2)";

        // Create a custom calculation engine that forces recalculation of the function
        MyEngine engine = new MyEngine();

        // Set calculation options with the custom engine (lifecycle rule: create)
        CalculationOptions opts = new CalculationOptions
        {
            CustomEngine = engine
        };

        // First calculation – the custom function is evaluated
        wb.CalculateFormula(opts);
        Console.WriteLine($"Initial B1 value: {ws.Cells["B1"].Value}");

        // Modify a dependent cell
        ws.Cells["A1"].PutValue(20);

        // Recalculate – because ForceRecalculate returns true, the cached result is invalidated
        wb.CalculateFormula(opts);
        Console.WriteLine($"After A1 change B1 value: {ws.Cells["B1"].Value}");

        // Save the workbook (lifecycle rule: save)
        wb.Save("InvalidateCacheDemo.xlsx");
    }

    // Custom engine that forces recalculation of the custom function
    class MyEngine : AbstractCalculationEngine
    {
        // ForceRecalculate method ensures the function is always recomputed
        public override bool ForceRecalculate(string functionName)
        {
            return functionName.Equals("MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase);
        }

        // Simple implementation: sum all numeric parameters
        public override void Calculate(CalculationData data)
        {
            if (data.FunctionName.Equals("MYVOLATILEFUNC", StringComparison.OrdinalIgnoreCase))
            {
                double sum = 0;
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object val = data.GetParamValue(i);
                    if (val is double d) sum += d;
                    else if (val is int iVal) sum += iVal;
                }
                data.CalculatedValue = sum;
            }
        }
    }
}
