// Title: Aspose.Cells .NET: Enable Iterative Calculation (Max 100 Iterations) and Resolve Circular References
// Description: Shows how to activate iterative calculation in Aspose.Cells, configure MaxIteration = 100 and MaxChange = 0.001, create a circular reference (A1 ↔ A2), execute workbook.CalculateFormula(), display the converged results, and save the file as IterativeCalculationDemo.xlsx.
// Keywords: Aspose.Cells iterative calculation | set MaxIteration Aspose.Cells | circular reference handling Aspose.Cells | MaxChange threshold Aspose.Cells | C# spreadsheet formula iteration
// Common Searches: enable iterative calculation Aspose.Cells .NET | Aspose.Cells max iteration example | handle circular reference formulas with Aspose.Cells | Aspose.Cells set MaxChange for convergence | C# iterative formula calculation Aspose.Cells
// Developer Intent: Turn on iterative calculation, define iteration limits, and evaluate circular‑reference formulas in a workbook using Aspose.Cells for .NET.
// Use Cases: Automate financial models that contain circular references by limiting iterations and ensuring convergence. | Prevent endless recalculation loops in engineering spreadsheets through MaxChange and iteration caps. | Create a sample workbook that demonstrates iterative convergence for teaching or testing purposes.
// AI Prompts: Generate C# code to run iterative calculation with a 50‑iteration limit and MaxChange of 0.0001 in Aspose.Cells. | Provide a snippet that reads the final values of cells involved in a circular reference after workbook.CalculateFormula(). | Explain how to disable iterative calculation and revert to standard formula evaluation in Aspose.Cells after processing.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to activate iterative calculation in Aspose.Cells, configure MaxIteration = 100 and MaxChange = 0.001, create a circular reference (A1 ↔ A2), execute workbook.CalculateFormula(), display the converged results, and save the file as IterativeCalculationDemo.xlsx.
    public class IterativeCalculationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Enable iterative calculation to resolve circular references
                workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

                // Set the maximum number of iterations
                workbook.Settings.FormulaSettings.MaxIteration = 100;

                // Define a maximum change threshold for convergence
                workbook.Settings.FormulaSettings.MaxChange = 0.001;

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create a circular reference: A1 depends on A2 and A2 depends on A1
                cells["A1"].Formula = "=A2+1";
                cells["A2"].Formula = "=A1+1";

                // Perform formula calculation (iterative mode will be used)
                workbook.CalculateFormula();

                // Output the calculated values
                Console.WriteLine($"A1 value after iterative calculation: {cells["A1"].Value}");
                Console.WriteLine($"A2 value after iterative calculation: {cells["A2"].Value}");

                // Save the workbook
                string outputPath = "IterativeCalculationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            IterativeCalculationDemo.Run();
        }
    }
}
