// Title: C# unit test with Aspose.Cells to validate a custom calculation engine that overrides SUM while preserving other functions
// Common Searches: aspnet cells custom calculation engine override sum function example | c# test to verify custom engine changes sum result in Aspose.Cells | how to keep average function unchanged when using custom calculation engine in Aspose.Cells | multiply sum result by factor using Aspose.Cells custom calculation engine | unit testing custom calculation engine with CalculationOptions in Aspose.Cells
// Tags: custom calculation engine sum override | Aspose.Cells CalculationOptions custom engine | C# unit test formula engine override | preserve default functions in custom engine | multiply sum result using custom engine

using System;
using Aspose.Cells;

namespace CustomEngineSumOverrideTest
{
    // Custom calculation engine that overrides the built‑in SUM function.
    // All other functions are left to the default calculation engine.
    // The example defines a SumOverrideEngine that inherits from AbstractCalculationEngine and intercepts the SUM function, multiplying its result by 10 while delegating all other functions to the default engine. A workbook is built with numbers and formulas for SUM and AVERAGE, default calculation captures expected values, then cells are cleared and recalculated using CalculationOptions with the custom engine. Assertions confirm that the SUM result is scaled (6 × 10 = 60) and the AVERAGE result remains unchanged (2). The workbook can be saved for visual inspection.
    public class SumOverrideEngine : AbstractCalculationEngine
    {
        // Multiply the normal SUM result by this factor to prove the override.
        private const double MultiplyFactor = 10.0;

        public override void Calculate(CalculationData data)
        {
            // Function name is case‑insensitive.
            string funcName = data.FunctionName?.Trim().ToUpperInvariant();

            if (funcName == "SUM")
            {
                double sum = 0.0;

                // Iterate over all parameters passed to SUM.
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);

                    // Parameter can be a single value or an array representing a range.
                    if (param is object[,] array)
                    {
                        // Walk through every cell in the array.
                        int rows = array.GetLength(0);
                        int cols = array.GetLength(1);
                        for (int r = 0; r < rows; r++)
                        {
                            for (int c = 0; c < cols; c++)
                            {
                                object cellVal = array.GetValue(r, c);
                                if (cellVal is double d)
                                    sum += d;
                                else if (cellVal is int iVal)
                                    sum += iVal;
                                else if (double.TryParse(cellVal?.ToString(), out double parsed))
                                    sum += parsed;
                            }
                        }
                    }
                    else
                    {
                        // Single scalar parameter.
                        if (param is double d)
                            sum += d;
                        else if (param is int iVal)
                            sum += iVal;
                        else if (double.TryParse(param?.ToString(), out double parsed))
                            sum += parsed;
                    }
                }

                // Apply the custom factor.
                data.CalculatedValue = sum * MultiplyFactor;
            }
            // For any other function we do NOT set CalculatedValue.
            // The default calculation engine will be used for those functions.
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Build a workbook with sample data and formulas.
                // ------------------------------------------------------------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // Input numbers.
                cells["A1"].PutValue(1);
                cells["A2"].PutValue(2);
                cells["A3"].PutValue(3);

                // Formula using built‑in SUM (will be overridden).
                cells["B1"].Formula = "=SUM(A1:A3)";

                // Formula using another built‑in function (should stay unchanged).
                cells["C1"].Formula = "=AVERAGE(A1:A3)";

                // ------------------------------------------------------------
                // 2. Calculate once with the default engine to capture expected values.
                // ------------------------------------------------------------
                wb.CalculateFormula();

                double expectedSum = cells["B1"].DoubleValue;      // Normal SUM = 6
                double expectedAvg = cells["C1"].DoubleValue;      // Normal AVERAGE = 2

                Console.WriteLine($"[Default] SUM result = {expectedSum}");
                Console.WriteLine($"[Default] AVERAGE result = {expectedAvg}");

                // ------------------------------------------------------------
                // 3. Reset the workbook (clear calculated values) and apply custom engine.
                // ------------------------------------------------------------
                // Clear previously calculated values.
                cells["B1"].PutValue(string.Empty);
                cells["C1"].PutValue(string.Empty);

                // Set up calculation options with our custom engine.
                CalculationOptions options = new CalculationOptions
                {
                    CustomEngine = new SumOverrideEngine()
                };

                // Re‑calculate using the custom engine.
                wb.CalculateFormula(options);

                double overriddenSum = cells["B1"].DoubleValue;   // Should be 6 * 10 = 60
                double unchangedAvg = cells["C1"].DoubleValue;    // Should still be 2

                Console.WriteLine($"[Custom] SUM result (overridden) = {overriddenSum}");
                Console.WriteLine($"[Custom] AVERAGE result (unchanged) = {unchangedAvg}");

                // ------------------------------------------------------------
                // 4. Simple test assertions.
                // ------------------------------------------------------------
                bool sumOverridden = Math.Abs(overriddenSum - expectedSum * 10.0) < 1e-9;
                bool avgUnchanged = Math.Abs(unchangedAvg - expectedAvg) < 1e-9;

                Console.WriteLine();
                Console.WriteLine("=== Test Results ===");
                Console.WriteLine($"SUM override test : {(sumOverridden ? "PASS" : "FAIL")}");
                Console.WriteLine($"Other function unchanged test : {(avgUnchanged ? "PASS" : "FAIL")}");

                // ------------------------------------------------------------
                // 5. Save the workbook for visual inspection (optional).
                // ------------------------------------------------------------
                try
                {
                    wb.Save("CustomEngineSumOverrideTest.xlsx");
                    Console.WriteLine("Workbook saved as CustomEngineSumOverrideTest.xlsx");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
