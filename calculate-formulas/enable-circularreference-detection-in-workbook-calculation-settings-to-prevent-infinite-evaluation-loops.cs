// Title: Enable circular reference detection with iterative calculation in Aspose.Cells using C#
// AI Prompts: Generate C# code that configures Aspose.Cells Workbook.Settings.FormulaSettings to turn on iterative calculation, set MaxIteration and MaxChange, and then calculate all formulas. | Provide a step‑by‑step guide for creating a circular reference between cells, enabling detection, and saving the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# enable formula iteration for circular references | Specify maximum iteration number and convergence threshold in Aspose.Cells formula settings | How to stop endless formula evaluation loops in an Excel workbook using Aspose.Cells | C# example for configuring formula settings to manage circular dependencies
// Tags: Aspose.Cells formula evaluation limits | circular dependency handling formula settings | configure iteration count and change tolerance Aspose.Cells | excel workbook formula calculation configuration | prevent infinite loops Aspose.Cells formulas

using System;
using Aspose.Cells;

// The example creates a workbook, introduces a circular reference between A1 and B1, enables iterative calculation with custom MaxIteration and MaxChange values to detect and resolve the loop, calculates all formulas, and saves the file as CircularReferenceDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set up a circular reference: A1 depends on B1 and B1 depends on A1
        sheet.Cells["A1"].Formula = "=B1+1";
        sheet.Cells["B1"].Formula = "=A1+1";

        // Enable iterative calculation to detect and resolve circular references
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
        workbook.Settings.FormulaSettings.MaxChange = 0.001;    // convergence threshold

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("CircularReferenceDemo.xlsx");
    }
}
