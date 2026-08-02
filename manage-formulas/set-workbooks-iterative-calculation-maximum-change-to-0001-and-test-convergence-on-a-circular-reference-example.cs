// Title: Aspose.Cells .NET: Set Iterative Calculation MaxChange to 0.001 and Test Circular Reference Convergence
// Description: Shows how to enable iterative calculation, set MaxIteration = 100 and MaxChange = 0.001, create a circular reference between A1 and B1, run CalculateFormula, display the converged values, and optionally save the workbook.
// Keywords: Aspose.Cells | .NET | C# | iterative calculation | MaxChange | MaxIteration | circular reference | formula convergence | EnableIterativeCalculation | CalculateFormula
// Common Searches: Aspose.Cells set MaxChange | iterative calculation circular reference .NET | how to limit change in Aspose.Cells formulas | C# Aspose.Cells iterative calculation example | test convergence of circular formulas Aspose.Cells
// Developer Intent: Demonstrate configuring iterative formula settings with a specific MaxChange and confirming convergence of a circular reference.
// Use Cases: Configure a workbook for iterative calculations with custom iteration limits. | Validate that circular formulas reach a stable value before exceeding MaxIteration. | Automate saving of the workbook after convergence for downstream processing. | Log final cell values and the MaxChange setting to aid debugging. | Measure performance by tracking iteration count and change thresholds.
// AI Prompts: Generate C# code that sets MaxChange to 0.0005 and runs iterative calculation on a different circular reference (e.g., C1 ↔ D1) using Aspose.Cells. | Write a function that returns true if the iterative calculation converged before reaching MaxIteration, and false otherwise. | Explain how to obtain the actual number of iterations performed after calling CalculateFormula in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to enable iterative calculation, set MaxIteration = 100 and MaxChange = 0.001, create a circular reference between A1 and B1, run CalculateFormula, display the converged values, and optionally save the workbook.
    public class IterativeCalculationMaxChangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Enable iterative calculation and set iteration parameters
                workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
                workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
                workbook.Settings.FormulaSettings.MaxChange = 0.001;   // maximum change threshold

                // Create a circular reference: A1 depends on B1 and B1 depends on A1
                worksheet.Cells["A1"].Formula = "=B1+1";
                worksheet.Cells["B1"].Formula = "=A1+1";

                // Perform calculation with the specified iterative settings
                workbook.CalculateFormula();

                // Output the results to verify convergence
                Console.WriteLine("A1 value after calculation: " + worksheet.Cells["A1"].Value);
                Console.WriteLine("B1 value after calculation: " + worksheet.Cells["B1"].Value);
                Console.WriteLine("MaxChange used: " + workbook.Settings.FormulaSettings.MaxChange);

                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save("IterativeCalculationMaxChangeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            IterativeCalculationMaxChangeDemo.Run();
        }
    }
}
