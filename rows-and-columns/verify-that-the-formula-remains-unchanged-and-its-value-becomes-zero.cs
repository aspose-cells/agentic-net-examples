// Title: Keep an Excel cell formula unchanged while forcing its displayed value to zero using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets a formula in a worksheet cell with Aspose.Cells and then assigns a custom result of 0 without altering the formula string. | Show how to use Aspose.Cells SetFormula method to preserve the original formula and override the cell's calculated value to zero, then output both the formula and the value. | Provide a C# example that verifies the Formula property remains "=SUM(1,2)" after calling SetFormula with a result of 0 in Aspose.Cells.
// Common Searches: Aspose.Cells C# set custom result for a formula cell while keeping the formula unchanged | how to force a cell value to zero without changing its formula using Aspose.Cells | SetFormula method custom result parameter example in Aspose.Cells .NET | verify cell formula remains the same after overriding calculated value in Aspose.Cells | saving workbook after setting formula with custom result in Aspose.Cells C#
// Tags: Aspose.Cells SetFormula custom result | preserve formula override value | force zero result Excel .NET | C# verify formula unchanged Aspose.Cells | save workbook after SetFormula Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, sets the formula "=SUM(1,2)" in cell A1 using SetFormula with a custom result of 0, prints the unchanged formula and the zero value for verification, and saves the file as FormulaZeroResult.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Target cell
        Cell cell = worksheet.Cells["A1"];

        // Define the formula we want to keep unchanged
        string formula = "=SUM(1,2)";

        // Set the formula and explicitly set its calculated result to zero
        // This keeps the formula text intact while the displayed value becomes 0
        cell.SetFormula(formula, 0);

        // Verify that the formula text is unchanged
        Console.WriteLine("Formula after SetFormula: " + cell.Formula); // Expected: =SUM(1,2)

        // Verify that the cell's value is now zero
        Console.WriteLine("Value after SetFormula: " + cell.Value); // Expected: 0

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("FormulaZeroResult.xlsx");
    }
}
