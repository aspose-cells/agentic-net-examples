// Title: Enable Iterative Calculation and Define Max Iterations for Circular Formulas in Aspose.Cells (.NET)
// Description: Creates a new Workbook, activates iterative calculation, sets maximum iteration count and change tolerance, demonstrates a circular reference (A1 ↔ A2), forces formula evaluation, prints the results, and saves the file as an .xlsx document.
// Keywords: Aspose.Cells iterative calculation | max iteration Aspose.Cells .NET | circular reference handling | formula convergence settings | MaxChange threshold | C# spreadsheet API | Excel circular formulas
// Common Searches: how to turn on iterative calculation in Aspose.Cells | set maximum iterations for circular references .NET | configure MaxChange for formula convergence Aspose.Cells | resolve circular formulas with Aspose.Cells C# | Aspose.Cells iterative mode example
// Developer Intent: Activate iterative mode and specify iteration limits to compute circular formulas reliably within a workbook.
// Use Cases: Financial models that contain self‑referencing cells need deterministic results. | Large engineering spreadsheets with looped calculations require controlled convergence precision. | Automated report generation where circular dependencies must be resolved without manual intervention.
// AI Prompts: Generate C# code that enables iterative calculation in Aspose.Cells, sets MaxIteration to 50 and MaxChange to 0.0001, then saves the workbook. | Explain how to detect non‑converging circular references using Aspose.Cells formula settings. | Provide performance tips for applying iterative calculation to workbooks with thousands of circular formulas.

using System;
using Aspose.Cells;

namespace AsposeCellsIterativeDemo
{
    // Creates a new Workbook, activates iterative calculation, sets maximum iteration count and change tolerance, demonstrates a circular reference (A1 ↔ A2), forces formula evaluation, prints the results, and saves the file as an .xlsx document.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Enable iterative calculation to resolve circular references
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

            // Set the maximum number of iterations for convergence
            workbook.Settings.FormulaSettings.MaxIteration = 20; // adjust as needed

            // Optional: set the maximum change threshold for convergence
            workbook.Settings.FormulaSettings.MaxChange = 0.001;

            // Define a circular reference scenario
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;
            cells["A1"].Formula = "=A2+1";
            cells["A2"].Formula = "=A1+1";

            // Calculate formulas with the iterative settings applied
            workbook.CalculateFormula();

            // Output the calculated values
            Console.WriteLine("A1 value: " + cells["A1"].Value);
            Console.WriteLine("A2 value: " + cells["A2"].Value);

            // Save the workbook (lifecycle save)
            workbook.Save("IterativeCalculationDemo.xlsx");
        }
    }
}
