// Title: Print the original formula and its zero result from cell E3 using Aspose.Cells for .NET (C#)
// Common Searches: how to retrieve and print a cell's original formula with Aspose.Cells in C# | Aspose.Cells C# display evaluated zero value of a specific Excel cell | read formula string and calculated value from cell E3 using Aspose.Cells .NET | console output of Excel cell formula and result with Aspose.Cells library
// Tags: read cell formula Aspose.Cells C# | evaluate cell value Aspose.Cells .NET | output Excel formula to console using Aspose.Cells | calculate workbook formulas before reading values Aspose.Cells | display zero result from SUM formula Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new workbook, assigns the formula "=SUM(A1:A2)" to cell E3, forces formula calculation, and writes both the unchanged formula string and its evaluated zero value to the console using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula in cell E3 that currently evaluates to zero
        // (A1 and A2 are empty, so SUM returns 0)
        Cell targetCell = sheet.Cells["E3"];
        targetCell.Formula = "=SUM(A1:A2)";

        // Calculate the formulas to ensure the value is up‑to‑date
        workbook.CalculateFormula();

        // Output the unchanged formula string
        Console.WriteLine("Formula in E3: " + targetCell.Formula);

        // Output the evaluated value (should be zero)
        Console.WriteLine("Value in E3: " + targetCell.Value);
    }
}
