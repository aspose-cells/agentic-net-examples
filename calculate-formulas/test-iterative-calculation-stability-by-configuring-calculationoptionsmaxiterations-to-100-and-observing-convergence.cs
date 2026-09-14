// Title: Enable iterative calculation with MaxIteration=100 and verify circular reference convergence using Aspose.Cells for .NET
// AI Prompts: Create a new Workbook, turn on iterative calculation, set MaxIteration to 100 and MaxChange to 0.001, define a circular reference (A1 = A2+1, A2 = A1+1), run CalculateFormula, and return the final values of A1 and A2. | Generate C# code that configures Aspose.Cells FormulaSettings for stability, limits the iteration count to 100, evaluates a circular reference, and prints the converged results.
// Common Searches: Aspose.Cells set maximum iterations for iterative formula calculation in C# | example of circular reference convergence with MaxIteration 100 using Aspose.Cells | how to test iterative calculation stability in Aspose.Cells .NET | C# code to limit iterative calculation loops to 100 and check result values
// Tags: maxiteration setting Aspose.Cells | circular reference convergence .NET | formula settings stability Aspose.Cells | limit calculation loops Aspose.Cells | formula convergence threshold C#

using System;
using Aspose.Cells;

namespace IterativeCalculationDemo
{
    // The sample creates a workbook, enables iterative calculation, sets MaxIteration to 100 and a convergence threshold of 0.001, defines a circular reference between A1 and A2, runs the calculation, prints the converged cell values, and saves the file as IterativeCalculationResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable iterative calculation and set maximum iterations to 100
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
            workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum iterations for circular references
            workbook.Settings.FormulaSettings.MaxChange = 0.001;   // convergence threshold

            // Set up a simple circular reference: A1 = A2 + 1, A2 = A1 + 1
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Formula = "=A2+1";
            sheet.Cells["A2"].Formula = "=A1+1";

            // Perform calculation
            workbook.CalculateFormula();

            // Output the resulting values to observe convergence
            Console.WriteLine("A1 value after iterative calculation: " + sheet.Cells["A1"].Value);
            Console.WriteLine("A2 value after iterative calculation: " + sheet.Cells["A2"].Value);

            // Optionally save the workbook to verify the results in Excel
            workbook.Save("IterativeCalculationResult.xlsx");
        }
    }
}
