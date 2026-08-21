// Title: Iterative Calculation with Custom MaxChange in Aspose.Cells for .NET
// Description: Shows how to activate iterative formula evaluation, set MaxIteration and a custom MaxChange tolerance, handle circular references, calculate cell values, and persist these settings in an .xlsx file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | iterative calculation | MaxChange | MaxIteration | circular reference | formula settings | Workbook.Settings | FormulaSettings | convergence threshold | calculate formula | save workbook
// Common Searches: Aspose.Cells enable iterative calculation | set MaxChange in Aspose.Cells .NET | circular reference calculation Aspose.Cells | configure formula iteration settings | limit iterations Aspose.Cells | iterative mode workbook settings
// Developer Intent: Turn on iterative mode and define convergence parameters (MaxIteration, MaxChange) to evaluate circular formulas in a workbook.
// Use Cases: Resolve circular references in financial models by applying iterative calculation with a specific tolerance. | Prevent endless recalculation loops in large spreadsheets by limiting iteration count and change magnitude. | Save and share workbooks that retain custom iterative settings for downstream users. | Adjust convergence thresholds programmatically based on runtime conditions.
// AI Prompts: Generate C# code that enables iterative calculation with MaxChange 0.0005 and MaxIteration 150 using Aspose.Cells. | Write a method that reads MaxChange and MaxIteration from appsettings.json and applies them before calling workbook.CalculateFormula(). | Explain how to verify that iterative settings are stored in the saved .xlsx file with Aspose.Cells APIs.

using System;
using Aspose.Cells;

// Shows how to activate iterative formula evaluation, set MaxIteration and a custom MaxChange tolerance, handle circular references, calculate cell values, and persist these settings in an .xlsx file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set up a circular reference for demonstration
        worksheet.Cells["A1"].Formula = "=B1+1";
        worksheet.Cells["B1"].Formula = "=A1+1";

        // Enable iterative calculation and configure convergence settings
        workbook.Settings.FormulaSettings.EnableIterativeCalculation = true; // Turn on iterative mode
        workbook.Settings.FormulaSettings.MaxIteration = 100;               // Maximum number of iterations
        workbook.Settings.FormulaSettings.MaxChange = 0.001;               // Custom maximum change threshold

        // Perform formula calculation using the configured settings
        workbook.CalculateFormula();

        // Output the calculated values
        Console.WriteLine("A1 value: " + worksheet.Cells["A1"].Value);
        Console.WriteLine("B1 value: " + worksheet.Cells["B1"].Value);
        Console.WriteLine("MaxChange used: " + workbook.Settings.FormulaSettings.MaxChange);

        // Save the workbook to verify settings are persisted
        workbook.Save("IterativeCalculationDemo.xlsx");
    }
}
