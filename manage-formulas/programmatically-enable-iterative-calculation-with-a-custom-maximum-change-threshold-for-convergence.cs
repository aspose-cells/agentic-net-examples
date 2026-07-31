// Title: Aspose.Cells for .NET – Enable Iterative Calculation with a Custom MaxChange
// Description: Demonstrates how to turn on iterative calculation in an Aspose.Cells workbook, set a custom convergence threshold (MaxChange) and iteration limit (MaxIteration), create a circular reference, run the calculation engine, display the final values, and save the file.
// Keywords: Aspose.Cells iterative calculation | C# MaxChange setting | custom convergence threshold .NET | circular reference formula Aspose | MaxIteration Aspose.Cells | Excel iterative mode programmatic | Aspose.Cells workbook settings
// Common Searches: how to enable iterative calculation in Aspose.Cells | set MaxChange and MaxIteration Aspose.Cells C# | circular reference handling with Aspose.Cells .NET | Aspose.Cells formula settings example | iterative calculation threshold Aspose.Cells
// Developer Intent: Configure a workbook to use iterative calculation with a specific MaxChange value and iteration cap, then evaluate circular formulas.
// Use Cases: Activate iterative mode and define convergence criteria before processing circular references. | Retrieve the computed values of cells involved in a loop after the engine stabilizes. | Persist the workbook with the applied iterative settings for downstream processing.
// AI Prompts: Generate C# code that enables iterative calculation in Aspose.Cells, sets MaxChange to 0.001, and runs a circular reference example. | Explain how Aspose.Cells determines convergence when MaxChange is configured and how to read the actual iteration count. | Show a step‑by‑step guide to adjust MaxIteration and MaxChange for large workbooks with circular formulas.

using System;
using Aspose.Cells;

namespace AsposeCellsIterativeDemo
{
    // Demonstrates how to turn on iterative calculation in an Aspose.Cells workbook, set a custom convergence threshold (MaxChange) and iteration limit (MaxIteration), create a circular reference, run the calculation engine, display the final values, and save the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (using the standard creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up iterative calculation with a custom maximum change threshold
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true; // enable iterative mode
            workbook.Settings.FormulaSettings.MaxIteration = 200;               // optional: limit iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.0005;               // custom convergence threshold

            // Create a circular reference to demonstrate iterative calculation
            cells["A1"].Formula = "=B1+1";
            cells["B1"].Formula = "=A1+1";

            // Perform the calculation using the workbook's calculation engine
            workbook.CalculateFormula();

            // Output the results after calculation
            Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].Value);
            Console.WriteLine("B1 value after iterative calculation: " + cells["B1"].Value);
            Console.WriteLine("Maximum change used: " + workbook.Settings.FormulaSettings.MaxChange);

            // Save the workbook (using the standard save rule)
            workbook.Save("IterativeCalculationDemo.xlsx");
        }
    }
}
