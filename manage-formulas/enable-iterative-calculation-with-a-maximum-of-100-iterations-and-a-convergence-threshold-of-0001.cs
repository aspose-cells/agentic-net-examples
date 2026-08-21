// Title: C# – Enable Iterative Calculation (100 iterations, 0.001 tolerance) in Aspose.Cells
// Description: Shows how to turn on iterative formula calculation in Aspose.Cells for .NET, set MaxIteration to 100 and MaxChange to 0.001, resolve circular references, evaluate the workbook, and optionally save the file.
// Keywords: Aspose.Cells | iterative calculation | MaxIteration | MaxChange | circular reference | C# example | FormulaSettings | Excel workbook | convergence tolerance | iteration limit
// Common Searches: Aspose.Cells enable iterative calculation | set MaxIteration Aspose.Cells C# | circular reference handling Aspose.Cells | configure MaxChange Aspose.Cells | iterative formula settings .NET | Aspose.Cells formula iteration limit
// Developer Intent: Configure Aspose.Cells to resolve circular references by enabling iterative calculation with a specific iteration count and convergence threshold.
// Use Cases: Financial models that contain circular formulas need controlled iteration to reach a stable result. | Automated Excel report generation where dependent cells must converge within a defined tolerance. | Data‑processing pipelines that recalculate worksheets repeatedly without risking infinite loops.
// AI Prompts: Generate C# code using Aspose.Cells to enable iterative calculation with 200 iterations and a 0.0005 tolerance. | Explain the impact of MaxIteration and MaxChange on circular reference resolution in Aspose.Cells and suggest optimal settings for large workbooks. | Create a sample workbook with several interdependent circular references and demonstrate how to configure iterative calculation for reliable convergence.

using System;
using Aspose.Cells;

// Shows how to turn on iterative formula calculation in Aspose.Cells for .NET, set MaxIteration to 100 and MaxChange to 0.001, resolve circular references, evaluate the workbook, and optionally save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable iterative calculation to resolve circular references
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

        // Set the maximum number of iterations to 100
        workbook.Settings.FormulaSettings.MaxIteration = 100;

        // Set the convergence threshold (maximum change) to 0.001
        workbook.Settings.FormulaSettings.MaxChange = 0.001;

        // Example circular reference for demonstration
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Formula = "=A2+1";
        sheet.Cells["A2"].Formula = "=A1+1";

        // Calculate all formulas using the configured iterative settings
        workbook.CalculateFormula();

        // Output the calculated values
        Console.WriteLine("A1 value: " + sheet.Cells["A1"].Value);
        Console.WriteLine("A2 value: " + sheet.Cells["A2"].Value);

        // Save the workbook (optional)
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}
