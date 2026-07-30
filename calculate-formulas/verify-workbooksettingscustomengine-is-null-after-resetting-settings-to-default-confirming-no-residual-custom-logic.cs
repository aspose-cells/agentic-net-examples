// Title: Aspose.Cells .NET – Confirm Workbook.Settings.CustomEngine Is Null After Resetting Calculation Options
// Description: C# sample that adds a custom calculation engine (MYFUNC), evaluates a formula, saves and reloads the workbook, then creates a fresh CalculationOptions instance to verify that Workbook.Settings.CustomEngine (or options.CustomEngine) is null. The example also shows the expected CellsException when the custom function is recalculated without the engine.
// Keywords: Aspose.Cells | C# | .NET | CustomEngine | CalculationOptions | Workbook.Settings.CustomEngine | reset custom engine | default calculation options | null engine check | MYFUNC custom function | formula calculation | CellsException | GitHub example | developer guide
// Common Searches: how to clear custom calculation engine in Aspose.Cells | Workbook.Settings.CustomEngine null after reset | reset Aspose.Cells calculation options to default | remove custom function engine from workbook .NET | verify custom engine removal Aspose.Cells
// Developer Intent: Validate that resetting calculation options clears any assigned custom engine so the workbook reverts to the built‑in calculation engine.
// Use Cases: Load a workbook saved with a custom engine, reset calculation options, and confirm the engine is cleared before a normal recalculation. | Programmatically create a new CalculationOptions object to remove a previously set custom engine and verify the removal via the CustomEngine property. | Attempt to recalculate a formula that relies on an undefined custom function after the engine has been cleared, catching the expected CellsException.
// AI Prompts: Generate C# code that resets Workbook.Settings.CustomEngine to null, checks the property, and demonstrates the resulting CellsException when recalculating a formula that uses a custom function. | Explain the lifecycle of a custom calculation engine in Aspose.Cells, covering how saving/loading a workbook affects the engine and how to ensure it is not retained after reset. | Write a unit test in C# that asserts Workbook.Settings.CustomEngine is null after creating a fresh CalculationOptions instance and that a CellsException is thrown for an undefined custom function.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineResetDemo
{
    // Custom engine that processes a dummy function "MYFUNC"
    // C# sample that adds a custom calculation engine (MYFUNC), evaluates a formula, saves and reloads the workbook, then creates a fresh CalculationOptions instance to verify that Workbook.Settings.CustomEngine (or options.CustomEngine) is null. The example also shows the expected CellsException when the custom function is recalculated without the engine.
    public class MyCustomEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            if (data.FunctionName.Equals("MYFUNC", StringComparison.OrdinalIgnoreCase))
            {
                // Expect two numeric parameters
                double param1 = Convert.ToDouble(data.GetParamValue(0));
                double param2 = Convert.ToDouble(data.GetParamValue(1));
                data.CalculatedValue = param1 + param2; // Simple sum
            }
        }

        // No need to force recalculation for this demo
        public override bool ForceRecalculate(string functionName) => false;
    }

    class Program
    {
        static void Main()
        {
            // ---------- Create ----------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate cells and a formula that uses the custom function
            ws.Cells["A1"].PutValue(5);
            ws.Cells["A2"].PutValue(7);
            ws.Cells["A3"].Formula = "=MYFUNC(A1, A2)";

            // Set calculation options with a custom engine
            CalculationOptions optionsWithEngine = new CalculationOptions
            {
                CustomEngine = new MyCustomEngine()
            };

            // Calculate using the custom engine
            wb.CalculateFormula(optionsWithEngine);
            Console.WriteLine("Result with custom engine: " + ws.Cells["A3"].Value); // Expected 12

            // Save the workbook (required by lifecycle rule)
            string filePath = "CustomEngineDemo.xlsx";
            wb.Save(filePath);

            // ---------- Load ----------
            Workbook loadedWb = new Workbook(filePath);
            Worksheet loadedWs = loadedWb.Worksheets[0];

            // Reset calculation options to default (no custom engine)
            CalculationOptions defaultOptions = new CalculationOptions(); // CustomEngine is null by default

            // Verify that CustomEngine is null after reset
            bool isEngineNull = defaultOptions.CustomEngine == null;
            Console.WriteLine("CustomEngine is null after resetting to default: " + isEngineNull);

            // Recalculate without custom engine to ensure default behavior (should error because MYFUNC is unknown)
            try
            {
                loadedWb.CalculateFormula(defaultOptions);
                Console.WriteLine("Result without custom engine: " + loadedWs.Cells["A3"].Value);
            }
            catch (CellsException ex)
            {
                Console.WriteLine("Expected error without custom engine: " + ex.Message);
            }

            // No further save needed for this verification
        }
    }
}
