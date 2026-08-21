// Title: Detect Circular References in Aspise.Cells with Iterative Calculation (C#)
// Description: This C# example shows how to create a workbook with a circular reference (A1 ↔ B1), enable iterative calculation via Workbook.Settings.FormulaSettings, set MaxIteration and MaxChange, run CalculateFormula, retrieve the converged values, and save the result to an XLSX file, preventing infinite evaluation loops.
// Keywords: Aspose.Cells circular reference | iterative calculation Aspose.Cells | EnableIterativeCalculation C# | MaxIteration MaxChange | prevent infinite formula loops | Aspose.Cells formula settings | C# workbook calculation | Aspose.Cells example
// Common Searches: Aspose.Cells enable iterative calculation C# | how to handle circular references in Aspose.Cells | set MaxIteration and MaxChange Aspose.Cells | prevent infinite formula loops Aspose.Cells .NET | Aspose.Cells circular reference detection example
// Developer Intent: Turn on iterative calculation and configure its limits to resolve circular references before evaluating formulas.
// Use Cases: Resolve mutually dependent cells (e.g., A1 and B1) without causing a stack overflow. | Process large workbooks that may contain hidden circular dependencies safely. | Generate reports where iterative formulas converge to a stable value and need to be saved.
// AI Prompts: Generate C# code that enables iterative calculation in Aspose.Cells with MaxIteration=200 and MaxChange=0.0001, then calculates all formulas. | Explain the algorithm Aspose.Cells uses to detect convergence when iterative calculation is enabled. | Create a unit test in C# that verifies A1 and B1 converge to expected values after iterative processing.

using System;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // This C# example shows how to create a workbook with a circular reference (A1 ↔ B1), enable iterative calculation via Workbook.Settings.FormulaSettings, set MaxIteration and MaxChange, run CalculateFormula, retrieve the converged values, and save the result to an XLSX file, preventing infinite evaluation loops.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up a circular reference for demonstration
            // A1 depends on B1 and B1 depends on A1
            cells["A1"].Formula = "=B1+1";
            cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation to detect and resolve circular references
            // This prevents infinite evaluation loops
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

            // Configure iterative calculation parameters
            workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.001;   // convergence threshold

            // Perform formula calculation
            workbook.CalculateFormula();

            // Output the calculated values after iterative processing
            Console.WriteLine("A1 value after calculation: " + cells["A1"].Value);
            Console.WriteLine("B1 value after calculation: " + cells["B1"].Value);

            // Save the workbook (lifecycle: save)
            workbook.Save("CircularReferenceResult.xlsx");
        }
    }
}
