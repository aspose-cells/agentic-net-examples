// Title: Aspose.Cells C# Example: Iterative Calculation with MaxIteration = 100
// Description: Shows how to enable iterative calculation in Aspose.Cells, configure MaxIteration to 100 and MaxChange to 0.001, set up a circular reference (A1 = A2+1, A2 = A1+1), execute workbook.CalculateFormula(), and read the resulting values.
// Keywords: Aspose.Cells | C# | iterative calculation | MaxIteration | MaxChange | circular reference | formula convergence | CalculateFormula | Excel automation | financial modeling
// Common Searches: enable iterative calculation Aspose.Cells C# | set MaxIteration to 100 in Aspose.Cells | circular reference convergence example Aspose.Cells | default MaxIteration value Aspose.Cells | how to limit formula iterations Aspose.Cells
// Developer Intent: Confirm that the workbook stops after 100 iterations and reaches the specified convergence threshold.
// Use Cases: Test stability of circular formulas in budgeting or engineering spreadsheets by adjusting MaxIteration and MaxChange. | Create automated regression tests that capture cell values after iterative calculation for comparison across library versions.
// AI Prompts: Write C# code that logs each cell's value after every iteration when using Aspose.Cells iterative calculation. | Explain how to fine‑tune FormulaSettings.MaxChange to obtain a desired precision in iterative formulas. | Provide a step‑by‑step guide to compare results before and after changing MaxIteration in a .NET workbook.

using System;
using Aspose.Cells;

namespace IterativeCalculationDemo
{
    // Shows how to enable iterative calculation in Aspose.Cells, configure MaxIteration to 100 and MaxChange to 0.001, set up a circular reference (A1 = A2+1, A2 = A1+1), execute workbook.CalculateFormula(), and read the resulting values.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable iterative calculation and set the maximum number of iterations to 100
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
            workbook.Settings.FormulaSettings.MaxIteration = 100;
            // Optional: define a convergence threshold
            workbook.Settings.FormulaSettings.MaxChange = 0.001;

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up a simple circular reference:
            // A1 = A2 + 1
            // A2 = A1 + 1
            cells["A1"].Formula = "=A2+1";
            cells["A2"].Formula = "=A1+1";

            // Perform calculation using the workbook's settings
            workbook.CalculateFormula();

            // Output the results to observe convergence after 100 iterations
            Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].Value);
            Console.WriteLine("A2 value after iterative calculation: " + cells["A2"].Value);

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("IterativeCalculationResult.xlsx");
        }
    }
}
