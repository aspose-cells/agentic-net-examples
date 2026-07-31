// Title: Aspose.Cells .NET: Enable Iterative Calculation, Create Circular Reference, Verify Convergence
// Description: Demonstrates how to turn on iterative calculation in Aspose.Cells, define a circular reference (A1 = B1+1, B1 = A1+1), set MaxIteration and MaxChange, run the calculation, and programmatically confirm that the result converges within the specified tolerance.
// Keywords: Aspose.Cells iterative calculation | circular reference .NET | MaxIteration Aspose.Cells | MaxChange tolerance | formula convergence C# | Aspose.Cells workbook settings | iterative formula evaluation
// Common Searches: how to enable iterative calculation in Aspose.Cells | Aspose.Cells circular reference example | verify formula convergence with Aspose.Cells | set MaxIteration and MaxChange Aspose.Cells .NET | iterative calculation sample code C#
// Developer Intent: Turn on iterative calculation, create a circular reference, and check that the formulas converge within a defined tolerance.
// Use Cases: Run financial models that contain mutually dependent cells by enabling iterative calculation and retrieving the stable values. | Control precision and performance of complex spreadsheets by adjusting MaxIteration and MaxChange settings. | Programmatically detect non‑convergent scenarios and trigger custom error handling or logging.
// AI Prompts: Generate C# code using Aspose.Cells to enable iterative calculation with custom MaxIteration and MaxChange, create a circular reference, and output whether the calculation converged. | Show how to log the iteration count and raise an exception when a circular reference does not converge in Aspose.Cells. | Explain how to read the final values and convergence status after iterative calculation in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsIterativeDemo
{
    // Demonstrates how to turn on iterative calculation in Aspose.Cells, define a circular reference (A1 = B1+1, B1 = A1+1), set MaxIteration and MaxChange, run the calculation, and programmatically confirm that the result converges within the specified tolerance.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Create a circular reference:
            // A1 = B1 + 1
            // B1 = A1 + 1
            cells["A1"].Formula = "=B1+1";
            cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation and set convergence parameters
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
            workbook.Settings.FormulaSettings.MaxIteration = 100;      // maximum number of iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.001;      // tolerance for convergence

            // Calculate formulas using the iterative settings
            workbook.CalculateFormula();

            // Retrieve the calculated values
            double a1Value = cells["A1"].DoubleValue;
            double b1Value = cells["B1"].DoubleValue;

            // Verify convergence: the absolute difference should be within the defined MaxChange
            double difference = Math.Abs(a1Value - b1Value);
            bool isConverged = difference <= workbook.Settings.FormulaSettings.MaxChange;

            // Output results
            Console.WriteLine($"A1 = {a1Value}");
            Console.WriteLine($"B1 = {b1Value}");
            Console.WriteLine($"Difference = {difference}");
            Console.WriteLine($"Converged within tolerance ({workbook.Settings.FormulaSettings.MaxChange}): {isConverged}");

            // Save the workbook (optional)
            workbook.Save("IterativeCircular.xlsx");
        }
    }
}
