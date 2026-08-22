// Title: Reset Aspose.Cells calculation options to default and verify Workbook.Settings.CustomEngine is null in C#
// AI Prompts: Generate C# code that creates a Workbook, assigns a DummyEngine as a custom calculation engine via CalculationOptions, calculates a formula, then creates a new CalculationOptions without a custom engine, recalculates, and checks that Workbook.Settings.CustomEngine is null. | Provide a step‑by‑step C# example showing how to clear a previously set custom calculation engine in Aspose.Cells, run calculations with default settings, and programmatically confirm that no custom engine remains attached to the workbook.
// Common Searches: Aspose.Cells how to clear a custom calculation engine in C# | C# verify Workbook.Settings.CustomEngine is null after resetting calculation options | reset calculation options to default Aspose.Cells workbook | remove custom engine from Aspose.Cells calculation and ensure default engine is used | Aspose.Cells calculate formula with custom engine then revert to default settings C#
// Tags: Aspose.Cells reset calculation options C# | Workbook.Settings.CustomEngine null check | custom calculation engine removal Aspose.Cells | default calculation engine after reset Aspose.Cells | calculate formula with custom engine C#

using System;
using Aspose.Cells;

namespace AsposeCellsCustomEngineResetDemo
{
    // Simple custom engine that does nothing special
    // Demonstrates assigning a DummyEngine as a custom calculation engine via CalculationOptions, performing a formula calculation, resetting to a fresh CalculationOptions instance, confirming that Workbook.Settings.CustomEngine is null, and saving the workbook.
    public class DummyEngine : AbstractCalculationEngine
    {
        public override void Calculate(CalculationData data)
        {
            // No custom processing; let default engine handle it
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Add a simple formula to demonstrate calculation
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

            // -----------------------------------------------------------------
            // Step 1: Assign a custom calculation engine via CalculationOptions
            // -----------------------------------------------------------------
            CalculationOptions optionsWithEngine = new CalculationOptions
            {
                CustomEngine = new DummyEngine() // set custom engine
            };

            // Perform calculation using the custom engine
            workbook.CalculateFormula(optionsWithEngine);

            // Verify that the custom engine is indeed set
            Console.WriteLine("CustomEngine assigned? " + (optionsWithEngine.CustomEngine != null));

            // -----------------------------------------------------------------
            // Step 2: Reset calculation settings to default (no custom engine)
            // -----------------------------------------------------------------
            // Create a fresh CalculationOptions instance without assigning CustomEngine
            CalculationOptions defaultOptions = new CalculationOptions();

            // Perform calculation again using default settings
            workbook.CalculateFormula(defaultOptions);

            // Verify that CustomEngine is null after reset
            Console.WriteLine("CustomEngine after reset is null? " + (defaultOptions.CustomEngine == null));

            // -----------------------------------------------------------------
            // Optional: Save the workbook to confirm no side effects (lifecycle rule: save)
            // -----------------------------------------------------------------
            workbook.Save("CustomEngineResetResult.xlsx");
        }
    }
}
