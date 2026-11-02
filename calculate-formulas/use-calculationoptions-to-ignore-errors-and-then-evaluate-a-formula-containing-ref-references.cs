// Title: How to ignore #REF! errors using CalculationOptions when evaluating formulas in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# calculate formula while ignoring #REF! errors | use CalculationOptions.IgnoreError to suppress reference errors during workbook calculation | evaluate Excel formula after deleting a column that creates #REF! with Aspose.Cells | skip #REF! errors when calling Workbook.CalculateFormula in .NET | configure Aspose.Cells to ignore errors in formula evaluation
// Tags: CalculationOptions.IgnoreError usage | handling #REF! during formula calculation | deleting column to generate #REF! in Aspose.Cells | Workbook.CalculateFormula with error suppression | C# Aspose.Cells formula evaluation options

using System;
using Aspose.Cells;

// The example creates a workbook, assigns a formula to A1, deletes column B to turn the B1 reference into a #REF! error, sets CalculationOptions.IgnoreError to true, calculates all formulas with Workbook.CalculateFormula, reads the resulting value from A1, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (using the provided create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Set up a formula that will later contain a #REF! error
        // Initially, A1 = B1 + C1 (both cells are empty, so result is 0)
        cells["A1"].Formula = "=B1+C1";

        // Delete column B to turn the reference to B1 into #REF!
        sheet.Cells.DeleteColumn(0); // Column B is index 1, after deletion of column A (index 0) it becomes index 0

        // At this point, the formula in A1 becomes something like =#REF!+C1

        // Configure calculation options to ignore errors like #REF!
        CalculationOptions calcOptions = new CalculationOptions();
        calcOptions.IgnoreError = true; // Instruct Aspose.Cells to ignore errors during calculation

        // Evaluate all formulas in the workbook using the configured options
        workbook.CalculateFormula(calcOptions);

        // Retrieve the evaluated value from A1
        var result = cells["A1"].Value;
        Console.WriteLine($"Evaluated result of A1 (ignoring errors): {result}");

        // Save the workbook (using the provided save rule)
        workbook.Save("ResultWorkbook.xlsx");
    }
}
