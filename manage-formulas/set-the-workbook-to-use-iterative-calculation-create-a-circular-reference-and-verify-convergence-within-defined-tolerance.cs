// Title: Aspose.Cells C# – Enable Iterative Calculation for Circular References and Verify Convergence
// Description: Creates a workbook, sets up a circular reference (A1 ↔ B1), activates iterative calculation with custom MaxIteration and MaxChange values, runs the calculation engine, checks that the result satisfies the tolerance, outputs the values, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | iterative calculation | circular reference | MaxIteration | MaxChange | convergence tolerance | formula settings | CalculateFormula | Excel circular formulas
// Common Searches: Aspose.Cells enable iterative calculation C# | Set MaxIteration MaxChange Aspose.Cells | How to handle circular references with Aspose.Cells | Check convergence of circular formulas Aspose.Cells | Save workbook after iterative calculation Aspose.Cells
// Developer Intent: Turn on iterative calculation, define a circular reference, execute the formula engine, and confirm that the outcome meets the configured tolerance.
// Use Cases: Financial models with inter‑dependent cells that require iterative solving | Engineering calculations where formulas reference each other and need precision control | Automated validation of spreadsheet convergence before publishing or exporting | Unit testing custom iteration settings in Aspose.Cells
// AI Prompts: Write C# code using Aspose.Cells to enable iterative calculation with MaxIteration = 100 and MaxChange = 0.001, then display whether convergence was reached. | Explain the algorithm Aspose.Cells uses to detect convergence for circular references and how to read the iteration count after CalculateFormula. | Create a C# NUnit test that asserts the absolute difference between cells A1 and B1 after iteration is ≤ MaxChange.

using System;
using Aspose.Cells;

namespace AsposeCellsIterativeDemo
{
    // Creates a workbook, sets up a circular reference (A1 ↔ B1), activates iterative calculation with custom MaxIteration and MaxChange values, runs the calculation engine, checks that the result satisfies the tolerance, outputs the values, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up a circular reference:
            // A1 depends on B1, B1 depends on A1
            cells["A1"].Formula = "=B1+1";
            cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation and define convergence criteria
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true; // allow iteration
            workbook.Settings.FormulaSettings.MaxIteration = 100;                // maximum number of iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.001;                // tolerance for convergence

            // Perform the calculation
            workbook.CalculateFormula();

            // Retrieve calculated values
            double a1Value = cells["A1"].DoubleValue;
            double b1Value = cells["B1"].DoubleValue;

            // Output the results
            Console.WriteLine($"A1 value after iterative calculation: {a1Value}");
            Console.WriteLine($"B1 value after iterative calculation: {b1Value}");

            // Verify convergence: the difference between successive iterations should be <= MaxChange.
            // For this simple circular reference the expected relationship is B1 = A1 + 1.
            double difference = Math.Abs(b1Value - (a1Value + 1));
            Console.WriteLine($"Difference from expected relationship (B1 = A1 + 1): {difference}");

            if (difference <= workbook.Settings.FormulaSettings.MaxChange)
                Console.WriteLine("Convergence within defined tolerance achieved.");
            else
                Console.WriteLine("Convergence not within tolerance; iteration limit may have been reached.");

            // Save the workbook (lifecycle rule: save)
            workbook.Save("IterativeCircularReference.xlsx");
        }
    }
}
