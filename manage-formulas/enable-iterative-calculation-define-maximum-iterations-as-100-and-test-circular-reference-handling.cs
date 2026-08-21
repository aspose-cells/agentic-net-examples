// Title: Aspose.Cells .NET: Enable Iterative Calculation (Max 100) and Resolve Circular References
// Description: Demonstrates how to turn on iterative calculation in Aspose.Cells, set MaxIteration to 100 (with a MaxChange threshold), create a circular reference between A1 and A2, calculate formulas, display the results, and save the workbook.
// Keywords: Aspose.Cells iterative calculation | max iteration 100 | circular reference handling | Aspose.Cells formula settings .NET | Enable iterative calculation C# | Aspose.Cells workbook.CalculateFormula | Excel circular reference Aspose
// Common Searches: how to enable iterative calculation in Aspose.Cells | set maximum iterations for formulas Aspose.Cells .NET | circular reference example Aspose.Cells C# | Aspose.Cells formula iteration limit | resolve circular formulas with Aspose.Cells
// Developer Intent: Configure Aspose.Cells to run iterative formula evaluation with a 100‑iteration cap and verify that circular references are processed safely.
// Use Cases: Financial models that contain self‑referencing cells need a bounded iterative solve. | Prevent endless calculation loops in large spreadsheets by defining MaxIteration and MaxChange. | Create a reproducible test workbook to benchmark iterative performance and result stability.
// AI Prompts: Show how to set MaxIteration to 200 and MaxChange to 0.0001 in Aspose.Cells .NET. | Provide code that logs each iteration’s cell values while resolving a circular reference with Aspose.Cells. | Explain how to disable iterative calculation after a workbook has been processed using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsIterativeDemo
{
    // Demonstrates how to turn on iterative calculation in Aspose.Cells, set MaxIteration to 100 (with a MaxChange threshold), create a circular reference between A1 and A2, calculate formulas, display the results, and save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Enable iterative calculation to resolve circular references
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

            // Set the maximum number of iterations (as requested)
            workbook.Settings.FormulaSettings.MaxIteration = 100;

            // Optional: define a small change threshold to stop iteration early
            workbook.Settings.FormulaSettings.MaxChange = 0.001;

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a circular reference:
            // A1 depends on A2, and A2 depends on A1
            cells["A1"].Formula = "=A2+1";
            cells["A2"].Formula = "=A1+1";

            // Perform formula calculation
            workbook.CalculateFormula();

            // Output the results after iterative calculation
            Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].Value);
            Console.WriteLine("A2 value after iterative calculation: " + cells["A2"].Value);

            // Save the workbook to verify the settings (lifecycle: save)
            workbook.Save("IterativeCircularReferenceDemo.xlsx");
        }
    }
}
