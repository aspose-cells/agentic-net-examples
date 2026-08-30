// Title: How to set Aspose.Cells FormulaSettings.MaxIteration to 100 for iterative calculations in C#
// AI Prompts: Generate C# code that enables iterative calculation and sets FormulaSettings.MaxIteration to 100 in an Aspose.Cells workbook. | Show an example that creates a workbook, applies a 100‑iteration limit for circular references, runs CalculateFormula, and saves the file. | Provide a snippet demonstrating how to configure Aspose.Cells to stop after 100 iterative calculation cycles.
// Common Searches: Aspose.Cells C# limit iterative formula recalculation to 100 cycles | Set MaxIteration property in Aspose.Cells FormulaSettings .NET | Enable iterative calculation with a maximum of 100 iterations using Aspose.Cells | How to prevent infinite loops in circular references with Aspose.Cells C# | Configure workbook calculation settings for iteration count in Aspose.Cells
// Tags: Aspose.Cells set MaxIteration C# | Iterative calculation limit Aspose.Cells | FormulaSettings EnableIterativeCalculation property | Circular reference handling Aspose.Cells | Workbook calculation settings Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new workbook, enables iterative calculation, sets MaxIteration to 100, adds a simple circular reference, calculates formulas, and saves the file as IterativeCalculationDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Enable iterative calculation and limit iterations to 100
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        workbook.Settings.FormulaSettings.MaxIteration = 100;

        // Example circular reference (optional, just to illustrate the setting)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].Formula = "=A2+1";
        sheet.Cells["A2"].Formula = "=A1+1";

        // Perform calculation with the specified settings
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}
