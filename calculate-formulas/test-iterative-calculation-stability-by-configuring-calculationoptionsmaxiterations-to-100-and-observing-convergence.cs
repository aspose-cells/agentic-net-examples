// Title: Aspose.Cells .NET: Test Iterative Calculation Stability with MaxIteration = 100
// Description: Demonstrates how to enable iterative calculation in Aspose.Cells, set MaxIteration to 100 and MaxChange to 0.001, create a circular reference (A1 = A2+1, A2 = A1+1), run the calculation engine, display the converged values, and optionally save the workbook.
// Keywords: Aspose.Cells iterative calculation | MaxIteration 100 | circular reference convergence | FormulaSettings MaxChange | C# Aspose.Cells example | calculate formulas .NET | iterative formula stability
// Common Searches: Aspose.Cells set MaxIteration 100 | iterative calculation example C# | how to test circular reference convergence Aspose.Cells | configure MaxChange for iterative formulas Aspose.Cells | iterative calculation stability .NET
// Developer Intent: Configure Aspose.Cells to run iterative formula evaluation with a limit of 100 iterations and verify that circular references converge.
// Use Cases: Validate iterative settings for financial models that contain circular dependencies. | Determine optimal MaxIteration and MaxChange values for engineering spreadsheets requiring stable convergence. | Generate a workbook that showcases iterative calculation results and persists the output for reporting.
// AI Prompts: Write C# code using Aspose.Cells to set MaxIteration to 200, log each iteration’s intermediate cell values, and output a convergence chart. | Explain the impact of MaxChange on iterative formula convergence in large workbooks and suggest best‑practice thresholds. | Create unit tests that assert the final values of A1 and A2 after iterative calculation with specific MaxIteration and MaxChange settings.

using System;
using Aspose.Cells;

namespace IterativeCalculationDemo
{
    // Demonstrates how to enable iterative calculation in Aspose.Cells, set MaxIteration to 100 and MaxChange to 0.001, create a circular reference (A1 = A2+1, A2 = A1+1), run the calculation engine, display the converged values, and optionally save the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable iterative calculation and set maximum iterations to 100
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
            workbook.Settings.FormulaSettings.MaxIteration = 100;
            // Optional: set a small MaxChange to define convergence threshold
            workbook.Settings.FormulaSettings.MaxChange = 0.001;

            // Create a simple circular reference:
            // A1 = A2 + 1
            // A2 = A1 + 1
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Formula = "=A2+1";
            sheet.Cells["A2"].Formula = "=A1+1";

            // Perform calculation
            workbook.CalculateFormula();

            // Output the resulting values to observe convergence
            Console.WriteLine("A1 value after iterative calculation: " + sheet.Cells["A1"].Value);
            Console.WriteLine("A2 value after iterative calculation: " + sheet.Cells["A2"].Value);

            // Save the workbook (optional, demonstrates create/save rule usage)
            workbook.Save("IterativeCalculationResult.xlsx");
        }
    }
}
