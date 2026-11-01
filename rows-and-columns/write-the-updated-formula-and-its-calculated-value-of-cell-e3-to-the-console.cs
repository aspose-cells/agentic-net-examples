// Title: Display the SUM formula and its computed value for cell E3 using Aspose.Cells for .NET
// AI Prompts: Write C# code that sets the formula "=SUM(A1:B1)" in cell E3 of an Aspose.Cells workbook, forces calculation, and prints the formula string and resulting value to the console. | Create a reusable C# method that takes a worksheet, a cell address, and a formula, applies the formula with Aspose.Cells, recalculates the workbook, and returns both the formula text and its evaluated result. | Adapt the example to log the formula and its evaluated result for any given cell into a JSON file instead of the console, using Aspose.Cells for .NET.
// Common Searches: how to get the formula string and calculated result of a specific cell with Aspose.Cells in C# | Aspose.Cells C# example printing cell formula and value to console | retrieve evaluated value after setting SUM formula in Aspose.Cells workbook | C# Aspose.Cells calculate formulas and read cell value programmatically
// Tags: Aspose.Cells set cell formula C# | Aspose.Cells calculate workbook formulas | Aspose.Cells read cell formula text | Aspose.Cells retrieve evaluated cell value | Aspose.Cells console output formula result

using System;
using Aspose.Cells;

// The program creates a new workbook, writes numbers to A1 and B1, assigns a SUM formula to cell E3, triggers formula recalculation, then reads both the formula text and its computed value from E3 and writes them to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for the formula
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);

        // Set a formula in cell E3 (e.g., sum of A1 and B1)
        Cell cellE3 = sheet.Cells["E3"];
        cellE3.Formula = "=SUM(A1:B1)";

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the formula text and its calculated value
        string formula = cellE3.Formula;
        object calculatedValue = cellE3.Value;

        // Output to console
        Console.WriteLine($"Formula in E3: {formula}");
        Console.WriteLine($"Calculated value: {calculatedValue}");
    }
}
