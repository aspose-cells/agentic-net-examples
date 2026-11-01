// Title: Print the original formula and its zero result from cell E3 using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a cell's Formula property and its evaluated numeric value, then writes both to the console with Aspose.Cells. | Show how to retrieve the unchanged formula string and the calculated zero result after calling Workbook.CalculateFormula in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells how to display a cell's formula and calculated value in console | C# retrieve original formula text after workbook.CalculateFormula | Get zero result of SUM formula from specific cell using Aspose.Cells | Read cell E3 formula and value with Aspose.Cells .NET example | Console output of formula and value for a worksheet cell in Aspose.Cells
// Tags: Aspose.Cells read cell formula | Aspose.Cells get calculated cell value | Aspose.Cells console output cell details | Aspose.Cells evaluate SUM formula | Aspose.Cells C# workbook calculation

using System;
using Aspose.Cells;

// The example creates a new workbook, assigns the formula "=SUM(A1:A2)" to cell E3, runs Workbook.CalculateFormula, then reads the unchanged Formula property and the evaluated numeric value (zero) and writes both to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula in cell E3 (row 2, column 4)
        // The formula SUM(A1:A2) will evaluate to 0 because A1 and A2 are empty
        Cell cell = sheet.Cells["E3"];
        cell.Formula = "=SUM(A1:A2)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the unchanged formula text
        string formula = cell.Formula;

        // Retrieve the calculated value (zero in this case)
        double value = cell.Value is double d ? d : 0.0;

        // Write the formula and its value to the console
        Console.WriteLine($"Formula: {formula}");
        Console.WriteLine($"Value: {value}");
    }
}
