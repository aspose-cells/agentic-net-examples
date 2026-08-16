// Title: Set ConvergenceThreshold (MaxChange) = 0.001 for Iterative Calculations in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to enable iterative calculation in Aspose.Cells, set the convergence threshold (MaxChange) to 0.001, optionally limit iterations, create a circular reference, calculate formulas, output results, and save the workbook.
// Keywords: Aspose.Cells iterative calculation | ConvergenceThreshold Aspose.Cells | MaxChange C# | Aspose.Cells circular reference | FormulaSettings precision | Aspose.Cells .NET example | Set MaxIteration Aspose.Cells | Iterative formula evaluation
// Common Searches: Aspose.Cells set convergence threshold | How to configure MaxChange in Aspose.Cells | Enable iterative calculation Aspose.Cells C# | Circular reference handling Aspose.Cells | Iterative formula settings Aspose.Cells .NET
// Developer Intent: Configure Aspose.Cells to use a 0.001 convergence threshold for iterative formula evaluation.
// Use Cases: Resolve circular references by enabling iterative calculation with a precise convergence threshold. | Balance performance and accuracy by adjusting MaxIteration and MaxChange values. | Validate iterative results by calculating formulas and saving the workbook.
// AI Prompts: Show C# code to set ConvergenceThreshold (MaxChange) to 0.001 in Aspose.Cells. | Provide an Aspose.Cells example that enables iterative calculation and defines MaxIteration and MaxChange. | Explain how MaxChange influences the resolution of circular references in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to enable iterative calculation in Aspose.Cells, set the convergence threshold (MaxChange) to 0.001, optionally limit iterations, create a circular reference, calculate formulas, output results, and save the workbook.
    class SetConvergenceThresholdDemo
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Enable iterative calculation and set related options
                workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
                workbook.Settings.FormulaSettings.MaxIteration = 100;          // optional: maximum iterations
                workbook.Settings.FormulaSettings.MaxChange = 0.001;          // precision for iterative calculations

                // Example circular reference to demonstrate iterative calculation
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].Formula = "=B1+1";
                sheet.Cells["B1"].Formula = "=A1+1";

                // Calculate all formulas
                workbook.CalculateFormula();

                // Output the results
                Console.WriteLine("A1 value: " + sheet.Cells["A1"].Value);
                Console.WriteLine("B1 value: " + sheet.Cells["B1"].Value);

                // Save the workbook
                workbook.Save("ConvergenceThresholdDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
