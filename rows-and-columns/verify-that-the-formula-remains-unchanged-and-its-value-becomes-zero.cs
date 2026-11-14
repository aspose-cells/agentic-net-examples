// Title: Verify that a cell's formula stays unchanged and its evaluated value is zero using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# verify formula string does not change after CalculateFormula | How to test if an Excel cell evaluates to zero after recalculation with Aspose.Cells | C# compare original and post‑calculation formula in Aspose.Cells workbook | Determine zero result of a formula without modifying the original expression in Aspose.Cells | Validate cell A1 formula integrity and zero value after workbook calculation in .NET
// Tags: preserve cell formula Aspose.Cells | recalculate workbook formulas Aspose.Cells | check cell value equals zero Aspose.Cells | compare original and updated formula strings Aspose.Cells | validate Excel cell calculation .NET

using System;
using Aspose.Cells;

// The example loads an Excel workbook, stores the original formula of cell A1, recalculates all formulas, then confirms that the formula text remains unchanged and that the resulting value is zero (handling both numeric and string representations). It outputs the verification results and saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Target cell that contains the formula to be tested
        Cell targetCell = sheet.Cells["A1"];

        // Preserve the original formula text
        string originalFormula = targetCell.Formula;

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the calculated value of the target cell
        object calculatedValue = targetCell.Value;

        // Verify that the formula text has not changed
        bool formulaUnchanged = string.Equals(originalFormula, targetCell.Formula, StringComparison.Ordinal);

        // Verify that the calculated value is zero (numeric zero or string "0")
        bool valueIsZero = false;
        if (calculatedValue is double d)
        {
            valueIsZero = Math.Abs(d) < 1e-12;
        }
        else if (calculatedValue is string s)
        {
            valueIsZero = s.Trim() == "0";
        }

        // Output verification results
        Console.WriteLine($"Formula unchanged: {formulaUnchanged}");
        Console.WriteLine($"Value is zero: {valueIsZero}");

        // Optionally, save the workbook if further processing is needed
        workbook.Save("output.xlsx");
    }
}
