// Title: Write C# unit tests to verify a custom Aspose.Cells calculation engine that doubles SUM results while leaving other functions unchanged
// AI Prompts: Generate an NUnit test method that confirms DoubleSumEngine returns twice the normal SUM value when ProcessBuiltInFunctions is true and leaves AVERAGE unchanged. | Create an xUnit test case that validates DoubleSumEngine with ProcessBuiltInFunctions set to false, ensuring SUM and AVERAGE produce their standard results. | Produce a MSTest suite that checks a custom AbstractCalculationEngine only overrides the SUM function and delegates all other built‑in functions to the default engine.
// Common Searches: how to unit test a custom Aspose.Cells calculation engine that overrides SUM in C# | Aspose.Cells DoubleSumEngine test ProcessBuiltInFunctions true false | C# example verifying custom calculation engine does not affect AVERAGE function | write automated tests for custom AbstractCalculationEngine in Aspose.Cells | validate custom SUM function behavior with Aspose.Cells CalculationOptions
// Tags: override SUM in Aspose.Cells custom engine | ProcessBuiltInFunctions property usage | Aspose.Cells calculation engine unit test | double result for SUM function | retain default behavior for other functions

using System;
using Aspose.Cells;

namespace CustomEngineSumTest
{
    // Custom calculation engine that optionally processes built‑in functions.
    // The sample defines DoubleSumEngine, a custom AbstractCalculationEngine that can optionally process built‑in functions. When enabled, it intercepts the SUM function, computes the sum of all arguments, doubles the result, and assigns it to the cell. All other functions, such as AVERAGE, are delegated to the default engine. The program runs two scenarios: one with ProcessBuiltInFunctions true (SUM result doubled, AVERAGE unchanged) and one with it false (both functions calculated normally), printing and asserting the expected outcomes.
    public class DoubleSumEngine : AbstractCalculationEngine
    {
        private readonly bool _processBuiltIn;

        public DoubleSumEngine(bool processBuiltIn)
        {
            _processBuiltIn = processBuiltIn;
        }

        // Enable or disable processing of built‑in functions.
        public override bool ProcessBuiltInFunctions => _processBuiltIn;

        // Only handle the SUM function when processing is enabled.
        public override void Calculate(CalculationData data)
        {
            if (data.FunctionName.Equals("SUM", StringComparison.OrdinalIgnoreCase) && ProcessBuiltInFunctions)
            {
                double sum = 0;
                // Iterate over all parameters (ranges, numbers, etc.).
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);

                    // If the parameter is a range (ReferredArea), sum each cell.
                    if (param is ReferredArea area)
                    {
                        for (int r = area.StartRow; r <= area.EndRow; r++)
                        {
                            for (int c = area.StartColumn; c <= area.EndColumn; c++)
                            {
                                object cellVal = area.GetValue(r, c);
                                if (cellVal != null && double.TryParse(cellVal.ToString(), out double d))
                                    sum += d;
                            }
                        }
                    }
                    // If the parameter is a scalar numeric value.
                    else if (param != null && double.TryParse(param.ToString(), out double d))
                    {
                        sum += d;
                    }
                }

                // Custom logic: double the normal SUM result.
                data.CalculatedValue = sum * 2;
            }
            // Other functions are left to the default engine.
        }

        // No need to force recalculation for any function.
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // ---------- Prepare workbook with test data ----------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Values for SUM and AVERAGE tests.
            ws.Cells["A1"].PutValue(1);
            ws.Cells["A2"].PutValue(2);
            ws.Cells["A3"].PutValue(3);
            ws.Cells["B1"].PutValue(10);
            ws.Cells["B2"].PutValue(20);
            ws.Cells["B3"].PutValue(30);

            // Formulas: one built‑in SUM, one built‑in AVERAGE.
            ws.Cells["C1"].Formula = "=SUM(A1:A3)";      // Expected normal sum = 6
            ws.Cells["C2"].Formula = "=AVERAGE(B1:B3)"; // Expected average = 20

            // ---------- Test 1: Engine processes built‑in functions ----------
            var engineProcessEnabled = new DoubleSumEngine(true);
            var optionsEnabled = new CalculationOptions { CustomEngine = engineProcessEnabled };
            wb.CalculateFormula(optionsEnabled);

            double sumResultEnabled = Convert.ToDouble(ws.Cells["C1"].Value);      // Should be 12 (6 * 2)
            double avgResultEnabled = Convert.ToDouble(ws.Cells["C2"].Value);      // Should remain 20

            Console.WriteLine("=== ProcessBuiltInFunctions = true ===");
            Console.WriteLine($"SUM result (doubled): {sumResultEnabled}"); // 12
            Console.WriteLine($"AVERAGE result (unchanged): {avgResultEnabled}"); // 20

            // Verify expectations.
            if (Math.Abs(sumResultEnabled - 12) > 0.0001)
                Console.WriteLine("ERROR: SUM was not doubled as expected.");
            if (Math.Abs(avgResultEnabled - 20) > 0.0001)
                Console.WriteLine("ERROR: AVERAGE should not be affected.");

            // ---------- Test 2: Engine does NOT process built‑in functions ----------
            // Reset workbook to force recalculation.
            ws.Cells["C1"].Formula = "=SUM(A1:A3)";
            ws.Cells["C2"].Formula = "=AVERAGE(B1:B3)";

            var engineProcessDisabled = new DoubleSumEngine(false);
            var optionsDisabled = new CalculationOptions { CustomEngine = engineProcessDisabled };
            wb.CalculateFormula(optionsDisabled);

            double sumResultDisabled = Convert.ToDouble(ws.Cells["C1"].Value); // Should be normal 6
            double avgResultDisabled = Convert.ToDouble(ws.Cells["C2"].Value); // Should be 20

            Console.WriteLine("\n=== ProcessBuiltInFunctions = false ===");
            Console.WriteLine($"SUM result (normal): {sumResultDisabled}"); // 6
            Console.WriteLine($"AVERAGE result (unchanged): {avgResultDisabled}"); // 20

            // Verify expectations.
            if (Math.Abs(sumResultDisabled - 6) > 0.0001)
                Console.WriteLine("ERROR: SUM should be calculated normally.");
            if (Math.Abs(avgResultDisabled - 20) > 0.0001)
                Console.WriteLine("ERROR: AVERAGE should remain unchanged.");

            // ---------- Save workbook (optional) ----------
            wb.Save("CustomEngineSumTest.xlsx");
        }
    }
}
