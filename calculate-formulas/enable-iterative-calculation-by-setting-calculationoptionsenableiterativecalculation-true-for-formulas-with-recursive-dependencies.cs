// Title: How to enable iterative calculation for circular references in Aspose.Cells using C#
// AI Prompts: Write C# code that configures Aspose.Cells to allow iterative calculation, defines iteration limit and tolerance, and evaluates circular formulas. | Show how to create a workbook where two cells depend on each other, enable iterative formula evaluation, and retrieve the computed values. | Demonstrate saving the workbook after iterative calculation with custom tolerance settings.
// Common Searches: Aspose.Cells C# enable iterative calculation for circular formulas | Set MaxIteration and MaxChange in Aspose.Cells formula settings .NET | Calculate workbook with circular reference using Aspose.Cells iterative mode | How to resolve circular dependencies in Excel files with Aspose.Cells C#
// Tags: iterative mode Aspose.Cells C# | circular formula handling Aspose.Cells | maxiteration maxchange configuration .NET | enable iterative formula evaluation workbook | recursive formula processing Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsIterativeCalculationDemo
{
    // Creates a workbook, defines a circular formula between A1 and B1, enables iterative calculation with custom MaxIteration and MaxChange values, calculates the formulas, outputs the results, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up a circular reference between A1 and B1
            cells["A1"].Formula = "=B1+1";
            cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation to resolve the circular reference
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
            // Optional: define iteration limits and tolerance
            workbook.Settings.FormulaSettings.MaxIteration = 100;
            workbook.Settings.FormulaSettings.MaxChange = 0.001;

            // Perform calculation
            workbook.CalculateFormula();

            // Output the results after iterative calculation
            Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].Value);
            Console.WriteLine("B1 value after iterative calculation: " + cells["B1"].Value);

            // Save the workbook (optional)
            workbook.Save("IterativeCalculationResult.xlsx");
        }
    }
}
