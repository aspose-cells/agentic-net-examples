// Title: Check that CalculationOptions.CustomEngine is null after resetting to defaults in Aspere.Cells for .NET
// Description: Demonstrates creating a workbook, assigning a dummy custom calculation engine via CalculationOptions, executing a formula, then instantiating a fresh CalculationOptions object and confirming its CustomEngine property is null, ensuring no residual custom logic before saving.
// Keywords: Aspose.Cells | .NET | C# | CalculationOptions | CustomEngine | reset to default | null check | custom calculation engine | formula calculation | Workbook settings
// Common Searches: Aspose.Cells reset CustomEngine | CalculationOptions default values .NET | How to clear custom calculation engine in Aspose.Cells | Check if CustomEngine is null after reset | Aspose.Cells C# verify custom engine removal
// Developer Intent: Validate that a newly created CalculationOptions instance has its CustomEngine property set to null, confirming that previous custom engine assignments are not retained.
// Use Cases: Ensure a clean calculation environment when switching between different custom engines. | Prevent unintended custom logic after reusing CalculationOptions across multiple workbook operations. | Unit‑test scenario to assert CustomEngine is cleared after resetting options. | Prepare a workbook for saving without lingering custom engine references.
// AI Prompts: Write C# code that creates a Workbook, sets a DummyEngine in CalculationOptions, runs CalculateFormula, then creates a new CalculationOptions and verifies CustomEngine is null. | Explain step‑by‑step how to reset Aspose.Cells CalculationOptions to default and confirm the CustomEngine property is cleared. | Generate an NUnit test that asserts CalculationOptions.CustomEngine is null after instantiating a fresh object. | Provide a brief guide on why resetting CustomEngine is important before saving a workbook in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineResetDemo
{
    // Simple custom calculation engine for demonstration
    // Demonstrates creating a workbook, assigning a dummy custom calculation engine via CalculationOptions, executing a formula, then instantiating a fresh CalculationOptions object and confirming its CustomEngine property is null, ensuring no residual custom logic before saving.
    public class DummyEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // No custom calculation; just let the default engine handle it
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Set a formula that could be processed by a custom engine
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // Create calculation options with a custom engine (lifecycle: create)
            CalculationOptions optionsWithEngine = new CalculationOptions
            {
                CustomEngine = new DummyEngine()
            };

            // Perform calculation using the custom engine
            workbook.CalculateFormula(optionsWithEngine);

            // Verify that the custom engine is set
            Console.WriteLine("CustomEngine set? " + (optionsWithEngine.CustomEngine != null));

            // Reset calculation options to default by creating a new instance
            CalculationOptions defaultOptions = new CalculationOptions();

            // Verify that after reset the CustomEngine property is null
            bool isEngineNull = defaultOptions.CustomEngine == null;
            Console.WriteLine("CustomEngine after reset is null? " + isEngineNull);

            // Save the workbook (lifecycle: save)
            workbook.Save("CustomEngineResetDemo.xlsx");
        }
    }
}
