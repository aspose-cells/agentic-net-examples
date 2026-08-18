// Title: Aspose.Cells .NET: Enable Iterative Calculation & Set Convergence for Circular References
// Description: Shows how to turn on iterative calculation in Aspose.Cells, configure MaxIteration and MaxChange, resolve a circular reference between A1 and B1, read the resulting values, and save the workbook.
// Keywords: Aspose.Cells iterative calculation | circular reference handling | MaxIteration | MaxChange | formula convergence | Workbook.Settings.FormulaSettings | C# Aspose.Cells example | calculate formulas | Excel circular dependency | Aspose.Cells .NET
// Common Searches: Aspose.Cells enable iterative calculation | set max iteration Aspose.Cells .NET | circular reference handling Aspose.Cells | configure convergence criteria Aspose.Cells | Aspose.Cells formula settings example | resolve circular formulas with Aspose.Cells
// Developer Intent: Activate iterative calculation and define iteration limits so that circular references are automatically resolved during formula evaluation.
// Use Cases: Break a simple A1 ↔ B1 circular dependency and obtain stable numeric results. | Apply iteration settings to a large workbook containing multiple circular formulas before bulk recalculation. | Generate a report with consistent values by saving the workbook after convergence is achieved. | Fine‑tune MaxIteration and MaxChange to balance performance and precision in financial models.
// AI Prompts: Write C# code using Aspose.Cells to enable iterative calculation with custom MaxIteration and MaxChange, then recalculate all formulas. | Explain the algorithm Aspose.Cells uses for circular references when iterative calculation is enabled and how to read the final cell values. | Suggest best‑practice values for MaxIteration and MaxChange to achieve reliable convergence in large spreadsheets.

using System;
using Aspose.Cells;

namespace AsposeCellsIterativeCalculationDemo
{
    // Shows how to turn on iterative calculation in Aspose.Cells, configure MaxIteration and MaxChange, resolve a circular reference between A1 and B1, read the resulting values, and save the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up a circular reference for demonstration
            // A1 depends on B1 and B1 depends on A1
            cells["A1"].Formula = "=B1+1";
            cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation to resolve the circular reference
            // and define convergence criteria (max iterations and max change)
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
            workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.001;   // convergence threshold

            // Perform formula calculation
            workbook.CalculateFormula();

            // Output the calculated values after iterative calculation
            Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].DoubleValue);
            Console.WriteLine("B1 value after iterative calculation: " + cells["B1"].DoubleValue);

            // Save the workbook (save rule)
            workbook.Save("IterativeCalculationResult.xlsx");
        }
    }
}
