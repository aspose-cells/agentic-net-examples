// Title: C# – Set Iterative Calculation (MaxIteration & MaxChange) for Circular References in Aspose.Cells
// Description: Demonstrates how to enable iterative calculation in Aspose.Cells, configure the maximum number of iterations and the convergence tolerance (MaxChange), resolve a circular reference between A1 and A2, retrieve the calculated values, and save the workbook.
// Keywords: Aspose.Cells | iterative calculation | MaxIteration | MaxChange | circular reference | C# | .NET | FormulaSettings | workbook calculation
// Common Searches: Aspose.Cells enable iterative calculation | set MaxIteration in Aspose.Cells C# | circular reference tolerance Aspose.Cells | FormulaSettings MaxChange example | resolve circular references Aspose.Cells .NET
// Developer Intent: Configure FormulaSettings to turn on iterative calculation, specify a maximum iteration count and a precision tolerance, and then evaluate circular formulas in a .NET workbook.
// Use Cases: Financial models that contain circular references and need deterministic iteration limits. | Engineering spreadsheets where convergence accuracy must be controlled via MaxChange. | Automated report generation that requires reliable results from inter‑dependent formulas.
// AI Prompts: Generate C# code using Aspose.Cells to enable iterative calculation with MaxIteration=100 and MaxChange=0.0001 for a workbook containing circular formulas. | Explain how Aspose.Cells determines convergence when MaxChange is set and outline troubleshooting steps for non‑converging circular references. | Create a unit test in C# that verifies the iterative calculation settings produce expected values for a simple circular reference between two cells.

using System;
using Aspose.Cells;

namespace AsposeCellsIterativeDemo
{
    // Demonstrates how to enable iterative calculation in Aspose.Cells, configure the maximum number of iterations and the convergence tolerance (MaxChange), resolve a circular reference between A1 and A2, retrieve the calculated values, and save the workbook.
    public class IterativeCalculationExample
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set up a circular reference: A1 depends on A2 and A2 depends on A1
                cells["A1"].Formula = "=A2+1";
                cells["A2"].Formula = "=A1+1";

                // Enable iterative calculation to resolve the circular reference
                workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

                // Set the maximum number of iterations (e.g., 50)
                workbook.Settings.FormulaSettings.MaxIteration = 50;

                // Set the precision tolerance (maximum change) for convergence (e.g., 0.001)
                workbook.Settings.FormulaSettings.MaxChange = 0.001;

                // Perform the calculation with the specified iterative settings
                workbook.CalculateFormula();

                // Output the calculated values
                Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].Value);
                Console.WriteLine("A2 value after iterative calculation: " + cells["A2"].Value);

                // Save the workbook to verify the settings (optional)
                workbook.Save("IterativeCalculationResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during iterative calculation: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            IterativeCalculationExample.Run();
        }
    }
}
