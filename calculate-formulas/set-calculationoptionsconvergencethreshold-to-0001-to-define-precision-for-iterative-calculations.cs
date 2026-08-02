// Title: Set Convergence Threshold (MaxChange) for Iterative Calculations in Aspose.Cells .NET
// Description: C# example that enables iterative calculation in Aspose.Cells, defines a circular reference, and sets Workbook.Settings.FormulaSettings.MaxChange to 0.001 to control precision and stop the loop when changes fall below the threshold.
// Keywords: Aspose.Cells iterative calculation | MaxChange convergence threshold | FormulaSettings EnableIterativeCalculation | circular reference precision .NET | Aspose.Cells calculation settings | set MaxChange 0.001
// Common Searches: Aspose.Cells set MaxChange | how to define convergence threshold in Aspose.Cells | iterative calculation example C# Aspose.Cells | circular reference handling Aspose.Cells | control precision of iterative formulas Aspose.Cells
// Developer Intent: Configure the convergence threshold to 0.001 so iterative formula evaluation stops when changes are smaller than this value.
// Use Cases: Terminate iterative calculations for financial models once the result stabilizes within 0.001, avoiding endless loops. | Improve performance of large spreadsheets by limiting iterations based on a defined precision level. | Ensure reproducible results in engineering simulations that rely on circular references and require specific accuracy.
// AI Prompts: Generate a C# snippet that sets EnableIterativeCalculation, MaxIteration, and MaxChange to 0.001 in Aspose.Cells and explains each setting. | Explain how FormulaSettings.MaxChange influences the stopping condition of iterative calculations in Aspose.Cells. | Create documentation that describes best practices for using convergence thresholds with circular references in Aspose.Cells.

using Aspose.Cells;
using System;

// C# example that enables iterative calculation in Aspose.Cells, defines a circular reference, and sets Workbook.Settings.FormulaSettings.MaxChange to 0.001 to control precision and stop the loop when changes fall below the threshold.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set up a circular reference for demonstration
            sheet.Cells["A1"].Formula = "=B1+1";
            sheet.Cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation and configure its parameters
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
            workbook.Settings.FormulaSettings.MaxIteration = 100;      // maximum number of iterations
            workbook.Settings.FormulaSettings.MaxChange = 0.001;      // convergence threshold (precision)

            // Perform calculation
            workbook.CalculateFormula();

            // Save the workbook
            string outputPath = "IterativeCalculationResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
