// Title: Aspose.Cells C# – Set Iterative Calculation MaxChange to 0.001 and Verify Circular Reference Convergence
// Description: This example shows how to enable iterative calculation in Aspose.Cells, configure MaxIteration (100) and MaxChange (0.001), create a circular reference (A1 = B1+1, B1 = A1+1), run CalculateFormula, and read the converged values. The workbook is then saved, demonstrating reliable convergence handling for inter‑dependent formulas.
// Keywords: Aspose.Cells iterative calculation | MaxChange 0.001 | circular reference handling .NET | FormulaSettings MaxIteration | C# spreadsheet convergence | Aspose.Cells calculate formula | iterative formula settings
// Common Searches: Aspose.Cells set MaxChange for iterative calculation | test circular reference convergence Aspose.Cells C# | enable iterative calculation Aspose.Cells .NET | maximum change threshold formula Aspose.Cells
// Developer Intent: Configure iterative calculation with a 0.001 change threshold, run the engine, and confirm that a circular reference reaches a stable result.
// Use Cases: Stabilize financial models that contain circular references by defining a precise MaxChange. | Control calculation performance in large workbooks by limiting iterations while ensuring convergence. | Automate export of converged spreadsheet data after iterative processing.
// AI Prompts: Write C# code using Aspose.Cells to enable iterative calculation, set MaxChange to 0.001, create a circular reference, calculate formulas, and display the final cell values. | Explain the impact of MaxIteration and MaxChange on formula convergence in Aspose.Cells and suggest optimal settings for typical business spreadsheets. | Provide a step‑by‑step tutorial for testing iterative calculation convergence in a .NET application, including how to read and verify the resulting values.

using System;
using Aspose.Cells;

namespace AsposeCellsIterativeCalculationDemo
{
    // This example shows how to enable iterative calculation in Aspose.Cells, configure MaxIteration (100) and MaxChange (0.001), create a circular reference (A1 = B1+1, B1 = A1+1), run CalculateFormula, and read the converged values. The workbook is then saved, demonstrating reliable convergence handling for inter‑dependent formulas.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Enable iterative calculation to resolve circular references
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

            // Set maximum number of iterations (optional, default is 100)
            workbook.Settings.FormulaSettings.MaxIteration = 100;

            // Set the maximum change threshold for convergence
            workbook.Settings.FormulaSettings.MaxChange = 0.001;

            // Create a circular reference:
            // A1 = B1 + 1
            // B1 = A1 + 1
            cells["A1"].Formula = "=B1+1";
            cells["B1"].Formula = "=A1+1";

            // Perform formula calculation with the iterative settings
            workbook.CalculateFormula();

            // Output the results to verify convergence
            Console.WriteLine("A1 value after calculation: " + cells["A1"].Value);
            Console.WriteLine("B1 value after calculation: " + cells["B1"].Value);
            Console.WriteLine("MaxChange used: " + workbook.Settings.FormulaSettings.MaxChange);

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            workbook.Save("IterativeCalculationDemo.xlsx");
        }
    }
}
