// Title: Print the original formula and its zero result from cell E3 using Aspose.Cells for .NET
// Description: Shows how to set the formula "=0" in cell E3 of a new workbook, trigger calculation, retrieve the unchanged formula string and the evaluated value, and write both to the console with Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET | get cell formula | read calculated value | workbook.CalculateFormula | console output | cell E3 | formula =0 | retrieve original formula | evaluate formula | Excel automation
// Common Searches: Aspose.Cells get original formula after calculation | C# read formula and value from cell | Aspose.Cells print cell formula to console | How to retrieve formula string in Aspose.Cells .NET | Calculate formulas Aspose.Cells and get result | Aspose.Cells example write formula and value
// Developer Intent: Output the unchanged formula and its evaluated zero value from cell E3 to the console.
// Use Cases: Debugging: log each cell's formula and result after workbook calculation. | Reporting: generate a text summary of formulas and their computed values before saving. | Automated testing: verify that specific cells return expected results such as zero.
// AI Prompts: Write C# code using Aspose.Cells that iterates over a range, prints each cell's original formula and calculated value to the console. | Provide an Aspose.Cells .NET example that reads a cell's formula, forces workbook calculation, and logs both the formula and formatted result to a file.

using System;
using Aspose.Cells;

// Shows how to set the formula "=0" in cell E3 of a new workbook, trigger calculation, retrieve the unchanged formula string and the evaluated value, and write both to the console with Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a formula in cell E3 that evaluates to zero
        worksheet.Cells["E3"].Formula = "=0";

        // Calculate formulas so the cell value is updated
        workbook.CalculateFormula();

        // Retrieve the unchanged formula and its calculated value
        string formula = worksheet.Cells["E3"].Formula;          // unchanged formula string
        string value   = worksheet.Cells["E3"].StringValue;     // formatted value, will be "0"

        // Output to console
        Console.WriteLine("Formula in E3: " + formula);
        Console.WriteLine("Value in E3: " + value);
    }
}
