// Title: Aspose.Cells .NET – Enable Iterative Calculation and Set MaxIteration/MaxChange for Circular References
// Description: Learn how to configure Aspose.Cells FormulaSettings to enable iterative calculation, define maximum iterations and precision tolerance, handle circular references (e.g., A1 ↔ B1), calculate the workbook, and save the result. Includes a complete C# example.
// Keywords: Aspose.Cells iterative calculation | EnableIterativeCalculation .NET | MaxIteration Aspose.Cells | MaxChange tolerance | circular reference Excel C# | FormulaSettings Aspose.Cells | Workbook.CalculateFormula | Aspose.Cells circular reference handling | Excel formula convergence | Aspose.Cells API example
// Common Searches: how to enable iterative calculation in Aspose.Cells .NET | set max iteration and tolerance for circular references Aspose.Cells | Aspose.Cells circular reference example C# | configure FormulaSettings for convergence in Aspose.Cells | Aspose.Cells iterative calculation settings tutorial
// Developer Intent: Configure FormulaSettings to activate iterative calculation, specify iteration limits and tolerance, and resolve circular references in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Resolve simple circular references (e.g., A1 = B1+1, B1 = A1+1) by enabling iterative calculation. | Stabilize large financial or engineering models with interdependent formulas through controlled iteration counts and precision. | Fine‑tune MaxChange to balance calculation speed against result accuracy before exporting the workbook.
// AI Prompts: Generate C# code that enables iterative calculation in Aspose.Cells, sets MaxIteration to 100 and MaxChange to 0.001, and evaluates a worksheet containing circular references. | Explain how Aspose.Cells handles circular references with iterative calculation and show how to read the final cell values after convergence. | Modify the provided example to use a different tolerance (e.g., 0.00001) and output the number of iterations performed. | Create a step‑by‑step guide for configuring FormulaSettings in Aspose.Cells to prevent #REF! errors caused by circular references.

using System;
using Aspose.Cells;

// Learn how to configure Aspose.Cells FormulaSettings to enable iterative calculation, define maximum iterations and precision tolerance, handle circular references (e.g., A1 ↔ B1), calculate the workbook, and save the result. Includes a complete C# example.
class IterativeCalculationDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable iterative calculation to handle circular references
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;

        // Set the maximum number of iterations (e.g., 50)
        workbook.Settings.FormulaSettings.MaxIteration = 50;

        // Set the precision tolerance (maximum change) for convergence
        workbook.Settings.FormulaSettings.MaxChange = 0.0001;

        // Define a simple circular reference:
        // A1 = B1 + 1
        // B1 = A1 + 1
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Formula = "=B1+1";
        sheet.Cells["B1"].Formula = "=A1+1";

        // Perform calculation using the iterative settings
        workbook.CalculateFormula();

        // Display the results after iterative calculation
        Console.WriteLine("A1 = " + sheet.Cells["A1"].Value);
        Console.WriteLine("B1 = " + sheet.Cells["B1"].Value);

        // Save the workbook (creation rule)
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}
