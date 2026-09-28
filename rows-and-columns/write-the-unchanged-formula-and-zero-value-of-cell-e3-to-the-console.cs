// Title: How to display the original "=0" formula and its evaluated zero value from cell E3 using Aspose.Cells for .NET
// AI Prompts: Provide C# code that assigns the formula "=0" to cell E3, forces workbook calculation, and prints both the stored formula string and the resulting numeric value to the console using Aspose.Cells. | Generate an Aspose.Cells snippet that reads the Formula property of a specific cell after calling CalculateFormula and writes the formula text together with the cell's Value to standard output. | Create a reusable C# method that takes a cell address, sets a zero‑producing formula, recalculates the workbook, and returns the original formula and its computed result for logging.
// Common Searches: Aspose.Cells .NET how to get cell formula after workbook.CalculateFormula | C# print Excel cell formula and calculated result using Aspose.Cells | retrieve unchanged formula string from worksheet cell after calculation Aspose.Cells | display zero result of =0 formula from cell E3 in console with Aspose.Cells
// Tags: read stored formula Aspose.Cells | calculate workbook formulas Aspose.Cells | output cell value to console C# | set zero-producing formula Aspose.Cells | retrieve original formula after calculation Aspose.Cells

using System;
using Aspose.Cells;

// The program creates a new workbook, assigns the "=0" formula to cell E3, calculates all formulas, then writes the unchanged formula text and its evaluated zero value to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a formula in cell E3 that evaluates to zero
        worksheet.Cells["E3"].Formula = "=0";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the unchanged formula text
        string unchangedFormula = worksheet.Cells["E3"].Formula;

        // Retrieve the calculated value (should be zero)
        object cellValue = worksheet.Cells["E3"].Value;

        // Output the formula and its zero value to the console
        Console.WriteLine("Formula in E3: " + unchangedFormula);
        Console.WriteLine("Value in E3: " + cellValue);
    }
}
