// Title: Aspose.Cells .NET: Enable Iterative Calculation to Handle Circular References
// Description: Demonstrates creating a workbook, inserting mutually dependent formulas, turning on iterative evaluation via FormulaSettings, optionally setting MaxIteration and MaxChange, executing CalculateFormula, and saving the result.
// Keywords: Aspose.Cells iterative calculation | circular reference handling .NET | FormulaSettings EnableIterativeCalculation | MaxIteration Aspose.Cells | MaxChange tolerance | C# spreadsheet calculation
// Common Searches: how to enable iterative calculation in Aspose.Cells | Aspose.Cells circular reference example | set iteration limits Aspose.Cells .NET | resolve circular formulas with Aspose.Cells | Aspose.Cells formula tolerance settings
// Developer Intent: Activate iterative mode so formulas that reference each other are evaluated without error.
// Use Cases: Financial models that require convergence of inter‑dependent cells | Simulation worksheets where values are refined until a tolerance is met | Automated generation of spreadsheets containing loan amortization or interest‑rate loops
// AI Prompts: Show C# code to turn on iterative calculation and define MaxIteration/MaxChange in Aspose.Cells. | Explain how Aspose.Cells processes circular references when iterative mode is enabled. | Provide a snippet that reads the iteration count after calling CalculateFormula.

using System;
using Aspose.Cells;

// Demonstrates creating a workbook, inserting mutually dependent formulas, turning on iterative evaluation via FormulaSettings, optionally setting MaxIteration and MaxChange, executing CalculateFormula, and saving the result.
class EnableIterativeCalculationDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define circular reference formulas
        sheet.Cells["A1"].Formula = "=B1+1";
        sheet.Cells["B1"].Formula = "=A1+1";

        // Enable iterative calculation to resolve the circular reference
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        // Optional: set iteration limits and tolerance
        workbook.Settings.FormulaSettings.MaxIteration = 100;
        workbook.Settings.FormulaSettings.MaxChange = 0.001;

        // Perform calculation
        workbook.CalculateFormula();

        // Display the calculated values
        Console.WriteLine("A1 value after iterative calculation: " + sheet.Cells["A1"].Value);
        Console.WriteLine("B1 value after iterative calculation: " + sheet.Cells["B1"].Value);

        // Save the workbook
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}
