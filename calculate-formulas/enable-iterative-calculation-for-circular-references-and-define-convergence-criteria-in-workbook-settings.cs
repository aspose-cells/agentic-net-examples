// Title: How to enable iterative calculation with convergence limits for circular references in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that configures Aspose.Cells to resolve circular references by turning on iterative calculation and specifying MaxIteration and MaxChange values. | Show how to set up a workbook with a circular formula dependency, enable iterative calculation, define convergence criteria, and retrieve the calculated results using Aspose.Cells for .NET. | Create a sample that creates an Excel file, applies iterative formula settings, runs CalculateFormula, and saves the file with the computed values.
// Common Searches: aspnet aspocells enable iterative calculation for circular references | set maxiteration and maxchange in Aspose.Cells formula settings c# example | how to calculate circular dependent cells with convergence threshold using Aspose.Cells | iterative formula calculation workbook settings Aspose.Cells .NET | c# Aspose.Cells circular reference resolution with iteration limit
// Tags: iterative calculation Aspose.Cells formula settings | circular reference resolution Aspose.Cells .NET | maxiteration maxchange convergence Aspose.Cells | calculate formulas with iteration limit C# | save workbook after iterative calculation Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, defines a circular reference between cells A1 and B1, enables iterative calculation, sets MaxIteration to 100 and MaxChange to 0.001 as convergence criteria, calculates the formulas, outputs the resulting values, and saves the workbook as IterativeCalculationDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define a circular reference: A1 depends on B1 and B1 depends on A1
        cells["A1"].Formula = "=B1+1";
        cells["B1"].Formula = "=A1+1";

        // Enable iterative calculation to resolve the circular reference
        // and set convergence criteria (max iterations and max change)
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
        workbook.Settings.FormulaSettings.MaxIteration = 100;   // maximum number of iterations
        workbook.Settings.FormulaSettings.MaxChange = 0.001;   // convergence threshold

        // Perform the calculation
        workbook.CalculateFormula();

        // Output the calculated values
        Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].Value);
        Console.WriteLine("B1 value after iterative calculation: " + cells["B1"].Value);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}
