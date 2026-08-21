// Title: C# unit test for Aspose.Cells custom engine that overrides SUM
// Description: Demonstrates how to create a test suite that validates a custom Aspose.Cells calculation engine. The engine intercepts the SUM function, doubles its result when ProcessBuiltInFunctions is true, and leaves all other functions (e.g., AVERAGE) to the default engine. Two workbooks are calculated—one with the custom SUM logic enabled and one without—to confirm expected outcomes.
// Keywords: Aspose.Cells | custom calculation engine | override SUM function | ProcessBuiltInFunctions | C# unit test | formula calculation | Excel automation | .NET testing
// Common Searches: Aspose.Cells custom engine unit test | override SUM in Aspose.Cells | ProcessBuiltInFunctions true example | C# test for custom calculation engine | verify custom SUM logic Aspose
// Developer Intent: Ensure the custom engine doubles SUM results when enabled while all other formulas compute normally.
// Use Cases: Run a test with ProcessBuiltInFunctions = true, calculate the workbook, and assert that SUM returns double the normal value and AVERAGE remains unchanged. | Run a test with ProcessBuiltInFunctions = false, calculate the workbook, and verify that SUM and other functions use the default calculation behavior. | Save the workbooks after each run for manual inspection or further automated validation.
// AI Prompts: Generate NUnit test methods that create workbooks, set SUM and AVERAGE formulas, apply CustomEngine with true/false flags, calculate formulas, and assert expected cell values. | Write an xUnit test class that verifies the custom SUM logic doubles the result when ProcessBuiltInFunctions is true and that other functions like AVERAGE use the default engine. | Provide MSTest code to compare outputs of the custom engine versus the default engine for SUM and AVERAGE, including workbook saving for manual review.

using System;
using Aspose.Cells;

// Demonstrates how to create a test suite that validates a custom Aspose.Cells calculation engine. The engine intercepts the SUM function, doubles its result when ProcessBuiltInFunctions is true, and leaves all other functions (e.g., AVERAGE) to the default engine. Two workbooks are calculated—one with the custom SUM logic enabled and one without—to confirm expected outcomes.
class CustomEngine : AbstractCalculationEngine
{
    private readonly bool _processBuiltIn;

    public CustomEngine(bool processBuiltIn)
    {
        _processBuiltIn = processBuiltIn;
    }

    // Enable processing of built‑in functions when true
    public override bool ProcessBuiltInFunctions => _processBuiltIn;

    public override void Calculate(CalculationData data)
    {
        // Custom handling for SUM when built‑in processing is enabled
        if (data.FunctionName.Equals("SUM", StringComparison.OrdinalIgnoreCase) && ProcessBuiltInFunctions)
        {
            double sum = 0;

            for (int i = 0; i < data.ParamCount; i++)
            {
                object param = data.GetParamValue(i);

                // Parameter can be a range (ReferredArea) or a scalar value
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
                else if (param != null && double.TryParse(param.ToString(), out double d))
                {
                    sum += d;
                }
            }

            // Custom logic: double the normal SUM result
            data.CalculatedValue = sum * 2;
        }
        // Other functions are left to the default engine
    }

    public override bool ForceRecalculate(string functionName) => false;
}

class Program
{
    static void Main()
    {
        // ------------------- Test 1: ProcessBuiltInFunctions = true -------------------
        Workbook wb1 = new Workbook();
        Worksheet ws1 = wb1.Worksheets[0];

        // Populate data
        ws1.Cells["A1"].PutValue(1);
        ws1.Cells["A2"].PutValue(2);
        ws1.Cells["A3"].PutValue(3);
        ws1.Cells["B1"].PutValue(10);
        ws1.Cells["B2"].PutValue(20);
        ws1.Cells["B3"].PutValue(30);

        // Formulas: SUM (to be overridden) and AVERAGE (should remain default)
        ws1.Cells["C1"].Formula = "=SUM(A1:A3)";
        ws1.Cells["C2"].Formula = "=AVERAGE(B1:B3)";

        // Calculate with custom engine that processes built‑in functions
        CalculationOptions optionsTrue = new CalculationOptions
        {
            CustomEngine = new CustomEngine(true)
        };
        wb1.CalculateFormula(optionsTrue);

        Console.WriteLine("=== ProcessBuiltInFunctions = true ===");
        Console.WriteLine($"SUM (custom, doubled)   : {ws1.Cells["C1"].Value}   // Expected: 12");
        Console.WriteLine($"AVERAGE (default)       : {ws1.Cells["C2"].Value}   // Expected: 20");

        // Save for manual verification if needed
        wb1.Save("CustomEngine_Sum_Doubled.xlsx");

        // ------------------- Test 2: ProcessBuiltInFunctions = false -------------------
        Workbook wb2 = new Workbook();
        Worksheet ws2 = wb2.Worksheets[0];

        // Same data as before
        ws2.Cells["A1"].PutValue(1);
        ws2.Cells["A2"].PutValue(2);
        ws2.Cells["A3"].PutValue(3);
        ws2.Cells["B1"].PutValue(10);
        ws2.Cells["B2"].PutValue(20);
        ws2.Cells["B3"].PutValue(30);

        // Same formulas
        ws2.Cells["C1"].Formula = "=SUM(A1:A3)";
        ws2.Cells["C2"].Formula = "=AVERAGE(B1:B3)";

        // Calculate with custom engine that does NOT process built‑in functions
        CalculationOptions optionsFalse = new CalculationOptions
        {
            CustomEngine = new CustomEngine(false)
        };
        wb2.CalculateFormula(optionsFalse);

        Console.WriteLine("\n=== ProcessBuiltInFunctions = false ===");
        Console.WriteLine($"SUM (default)           : {ws2.Cells["C1"].Value}   // Expected: 6");
        Console.WriteLine($"AVERAGE (default)       : {ws2.Cells["C2"].Value}   // Expected: 20");

        // Save for manual verification if needed
        wb2.Save("CustomEngine_Sum_Default.xlsx");
    }
}
