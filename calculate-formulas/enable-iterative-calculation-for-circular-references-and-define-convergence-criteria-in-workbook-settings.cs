// Title: Aspose.Cells C# – Enable Iterative Calculation & Set Convergence for Circular References
// Description: Demonstrates how to activate iterative calculation in an Aspose.Cells workbook, define MaxIteration and MaxChange thresholds, evaluate circular references between cells, retrieve the final values, and save the file.
// Keywords: Aspose.Cells iterative calculation | circular reference handling | MaxIteration | MaxChange | formula settings | C# workbook | EnableIterativeCalculation | Aspose.Cells .NET | Excel circular reference | iterative formula calculation
// Common Searches: How to enable iterative calculation in Aspose.Cells C# | Aspose.Cells set MaxIteration and MaxChange | Circular reference handling with Aspose.Cells .NET | Calculate formulas using iterative mode Aspose.Cells | Save workbook after iterative calculation Aspose.Cells
// Developer Intent: Turn on iterative calculation and define convergence limits to resolve circular references in an Aspose.Cells workbook.
// Use Cases: Fix circular references in financial models by configuring iterative settings and tolerance values. | Run stable engineering calculations on self‑referencing cells without manual intervention. | Programmatically generate a spreadsheet, apply iterative parameters, compute final results, and persist the workbook.
// AI Prompts: Show me C# code to enable iterative calculation with custom MaxIteration and MaxChange in Aspose.Cells. | Provide an example that handles circular references, runs calculation, and saves the workbook using Aspose.Cells. | Explain how to read the final cell values after Aspose.Cells completes iterative calculation.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to activate iterative calculation in an Aspose.Cells workbook, define MaxIteration and MaxChange thresholds, evaluate circular references between cells, retrieve the final values, and save the file.
    public class IterativeCalculationDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define a circular reference for demonstration
                worksheet.Cells["A1"].Formula = "=B1+1";
                worksheet.Cells["B1"].Formula = "=A1+1";

                // Enable iterative calculation and set convergence criteria
                workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
                workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
                workbook.Settings.FormulaSettings.MaxChange = 0.001;   // minimum change threshold for convergence

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Output the results after calculation
                Console.WriteLine("A1 value after iterative calculation: " + worksheet.Cells["A1"].DoubleValue);
                Console.WriteLine("B1 value after iterative calculation: " + worksheet.Cells["B1"].DoubleValue);

                // Save the workbook
                string outputPath = "IterativeCalculationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error during iterative calculation demo: " + ex.Message);
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
