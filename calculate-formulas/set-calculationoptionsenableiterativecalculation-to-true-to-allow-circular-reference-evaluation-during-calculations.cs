// Title: How to enable iterative calculation for circular references in Aspose.Cells using C#
// AI Prompts: Generate C# code that activates iterative calculation in Aspose.Cells and configures MaxIteration and MaxChange. | Show how to create a circular reference in a worksheet and then calculate formulas with iteration enabled. | Provide a snippet that reads the resulting cell values after iterative formula evaluation and saves the workbook.
// Common Searches: Aspose.Cells enable iterative calculation to resolve circular references in .NET | C# set MaxIteration and MaxChange for formula calculation with Aspose.Cells | How to calculate circular reference formulas using Aspose.Cells workbook | Iterative formula evaluation example Aspose.Cells C# | Enable iterative calculation settings in Aspose.Cells workbook programmatically
// Tags: iterative calculation Aspose.Cells C# | circular reference handling Aspose.Cells | set MaxIteration MaxChange Aspose.Cells | formula calculation settings Aspose.Cells | save workbook after iterative calculation Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsIterativeCalculationDemo
{
    // The example creates a workbook, defines a circular reference between cells A1 and B1, enables iterative calculation with custom MaxIteration and MaxChange limits, calculates the formulas, outputs the computed values, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set up a circular reference for demonstration
            cells["A1"].Formula = "=B1+1";
            cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation to resolve the circular reference
            workbook.Settings.FormulaSettings.EnableIterativeCalculation = true;
            // Optional: define iteration limits
            workbook.Settings.FormulaSettings.MaxIteration = 100;
            workbook.Settings.FormulaSettings.MaxChange = 0.001;

            // Perform calculation
            workbook.CalculateFormula();

            // Output the calculated values
            Console.WriteLine("A1 value after iterative calculation: " + cells["A1"].DoubleValue);
            Console.WriteLine("B1 value after iterative calculation: " + cells["B1"].DoubleValue);

            // Save the workbook (optional)
            workbook.Save("IterativeCalculationDemo.xlsx");
        }
    }
}
