// Title: Use Worksheet.CalculateFormula with custom CalculationOptions to evaluate an Excel formula string in C# (Aspose.Cells)
// Common Searches: how to evaluate an Excel formula string without writing it to a cell using Aspose.Cells C# | C# Aspose.Cells Worksheet.CalculateFormula with custom CalculationOptions example | evaluate SUM(1,2,3) using Aspose.Cells calculation options ignore errors | run a formula programmatically in Aspose.Cells without adding it to the worksheet | set CalculationOptions.IgnoreError while calculating a formula with Aspose.Cells
// Tags: custom CalculationOptions for Worksheet.CalculateFormula | programmatic Excel formula evaluation in Aspose.Cells | ignore errors during formula calculation Aspose.Cells | evaluate formula without inserting into worksheet | C# Aspose.Cells SUM function calculation

using Aspose.Cells;
using System;

// The example creates an in‑memory workbook, obtains the first worksheet, defines a formula string (SUM(1,2,3)), configures a CalculationOptions object (e.g., IgnoreError = true), and calls Worksheet.CalculateFormula with the formula and options to compute the result without placing the formula in a cell, then prints the evaluated value.
class Program
{
    static void Main()
    {
        // Create a new workbook in memory
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Formula to evaluate (does not need to be placed in a cell)
        string formula = "SUM(1, 2, 3)";

        // Create custom calculation options
        CalculationOptions calcOptions = new CalculationOptions();
        // Example option: ignore errors during calculation
        calcOptions.IgnoreError = true;
        // Additional options can be set here as needed, e.g.,
        // calcOptions.EnableIterativeCalculation = true;

        // Evaluate the formula using the worksheet's CalculateFormula method
        object result = sheet.CalculateFormula(formula, calcOptions);

        // Display the result
        Console.WriteLine($"Result of formula '{formula}' is: {result}");
    }
}
