// Title: Unit test suite to verify a custom Aspose.Cells .NET engine overrides SUM while keeping other functions unchanged
// Description: Shows a CustomSumEngine (inherits AbstractCalculationEngine) that optionally processes built‑in functions. When enabled it intercepts SUM, doubles the result, and lets AVERAGE and all other functions fall back to the default engine. The test workbook fills A1:A3, adds =SUM and =AVERAGE, runs calculations with ProcessBuiltInFunctions true and false, and validates the outcomes.
// Keywords: Aspose.Cells | .NET | C# | CustomCalculationEngine | override SUM | process built‑in functions | unit test | formula calculation | AbstractCalculationEngine | Excel automation | test suite
// Common Searches: Aspose.Cells custom engine override SUM example | How to test custom calculation engine in Aspose.Cells .NET | ProcessBuiltInFunctions true SUM double result | Unit testing Aspose.Cells formula engine | CustomSumEngine Aspose.Cells tutorial
// Developer Intent: Validate that a custom calculation engine can replace the SUM function while leaving all other Excel functions untouched.
// Use Cases: Run workbook calculation with CustomSumEngine(true) to ensure SUM returns twice the normal total and AVERAGE remains standard. | Run the same workbook with CustomSumEngine(false) to confirm SUM uses the default implementation and other functions are unaffected. | Integrate the test into CI pipelines to catch regressions in custom engine behavior.
// AI Prompts: Generate an NUnit test that creates a workbook, applies =SUM and =AVERAGE, calculates with CustomSumEngine(true), and asserts B1 = 12 and B2 = 2. | Write a MSTest method that runs the workbook with CustomSumEngine(false) and verifies B1 = 6 while B2 stays 2. | Provide a xUnit test suite that programmatically builds the sample workbook, toggles ProcessBuiltInFunctions, executes calculations, and checks expected cell values.

using System;
using Aspose.Cells;

namespace CustomEngineSumTest
{
    // Custom calculation engine that can optionally process built‑in functions.
    // Shows a CustomSumEngine (inherits AbstractCalculationEngine) that optionally processes built‑in functions. When enabled it intercepts SUM, doubles the result, and lets AVERAGE and all other functions fall back to the default engine. The test workbook fills A1:A3, adds =SUM and =AVERAGE, runs calculations with ProcessBuiltInFunctions true and false, and validates the outcomes.
    public class CustomSumEngine : AbstractCalculationEngine
    {
        private readonly bool _processBuiltIn;

        public CustomSumEngine(bool processBuiltIn)
        {
            _processBuiltIn = processBuiltIn;
        }

        // Enable or disable processing of built‑in functions.
        public override bool ProcessBuiltInFunctions => _processBuiltIn;

        // Only handle the SUM function when processing built‑in functions is enabled.
        public override void Calculate(CalculationData data)
        {
            if (data.FunctionName.Equals("SUM", StringComparison.OrdinalIgnoreCase) && ProcessBuiltInFunctions)
            {
                double sum = 0;
                // Iterate over all parameters (ranges, numbers, etc.).
                for (int i = 0; i < data.ParamCount; i++)
                {
                    object param = data.GetParamValue(i);
                    // If the parameter is a range, iterate its cells.
                    if (param is ReferredArea area)
                    {
                        for (int r = area.StartRow; r <= area.EndRow; r++)
                        {
                            for (int c = area.StartColumn; c <= area.EndColumn; c++)
                            {
                                object cellVal = area.GetValue(r, c);
                                if (cellVal != null)
                                    sum += Convert.ToDouble(cellVal);
                            }
                        }
                    }
                    else if (param != null) // Single numeric argument.
                    {
                        sum += Convert.ToDouble(param);
                    }
                }
                // Custom logic: double the normal SUM result.
                data.CalculatedValue = sum * 2;
            }
            // All other functions fall back to the default engine automatically.
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

            // Fill cells A1:A3 with 1, 2, 3.
            ws.Cells["A1"].PutValue(1);
            ws.Cells["A2"].PutValue(2);
            ws.Cells["A3"].PutValue(3);

            // Formula that uses SUM (will be overridden) and AVERAGE (should stay default).
            ws.Cells["B1"].Formula = "=SUM(A1:A3)";
            ws.Cells["B2"].Formula = "=AVERAGE(A1:A3)";

            // ---------- Test 1: Engine processes built‑in functions ----------
            var engineProcessEnabled = new CustomSumEngine(true);
            var optionsEnabled = new CalculationOptions { CustomEngine = engineProcessEnabled };
            wb.CalculateFormula(optionsEnabled);

            Console.WriteLine("=== Engine with ProcessBuiltInFunctions = true ===");
            Console.WriteLine($"B1 (SUM)   Expected: 12   Actual: {ws.Cells["B1"].Value}");
            Console.WriteLine($"B2 (AVERAGE) Expected: 2    Actual: {ws.Cells["B2"].Value}");

            // ---------- Test 2: Engine does NOT process built‑in functions ----------
            // Reset formulas to force recalculation.
            ws.Cells["B1"].Formula = "=SUM(A1:A3)";
            ws.Cells["B2"].Formula = "=AVERAGE(A1:A3)";

            var engineProcessDisabled = new CustomSumEngine(false);
            var optionsDisabled = new CalculationOptions { CustomEngine = engineProcessDisabled };
            wb.CalculateFormula(optionsDisabled);

            Console.WriteLine("\n=== Engine with ProcessBuiltInFunctions = false ===");
            Console.WriteLine($"B1 (SUM)   Expected: 6    Actual: {ws.Cells["B1"].Value}");
            Console.WriteLine($"B2 (AVERAGE) Expected: 2    Actual: {ws.Cells["B2"].Value}");

            // ---------- Save workbook (optional) ----------
            wb.Save("CustomEngineSumTest.xlsx");
        }
    }
}
