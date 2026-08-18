// Title: Aspose.Cells for .NET (C#): Display updated formula and calculated value of cell E3
// Description: Creates a workbook, assigns numeric values to B3, C3, and D3, sets a SUM formula in E3, forces calculation with Workbook.CalculateFormula, and writes both the formula string and the evaluated result of E3 to the console.
// Keywords: Aspose.Cells | .NET | C# | cell formula | CalculateFormula | SUM function | console output | read cell formula | retrieve cell value | Excel automation
// Common Searches: Aspose.Cells get formula after calculation | C# display cell formula and value | How to read Excel formula with Aspose.Cells | Calculate and print cell result in .NET
// Developer Intent: Print the formula assigned to E3 and its evaluated numeric result.
// Use Cases: Debugging: verify that a programmatically set formula returns the expected total. | Reporting: include both the original Excel expression and its computed value in logs or output files. | Testing: automate validation of formula logic by comparing printed results with expected sums.
// AI Prompts: Generate C# code using Aspose.Cells that sets a SUM formula in a cell, calculates the workbook, and prints the formula and its value. | Explain how to access the Formula and Value properties of a cell after calling Workbook.CalculateFormula in Aspose.Cells. | Show a formatted console output that clearly distinguishes the formula string from the numeric result for a given cell.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    // Creates a workbook, assigns numeric values to B3, C3, and D3, sets a SUM formula in E3, forces calculation with Workbook.CalculateFormula, and writes both the formula string and the evaluated result of E3 to the console.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate cells that will be used in the formula
            cells["B3"].PutValue(10);
            cells["C3"].PutValue(20);
            cells["D3"].PutValue(30);

            // Set a formula in cell E3 that sums B3, C3 and D3
            cells["E3"].Formula = "=SUM(B3:D3)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the updated formula and its calculated value
            Console.WriteLine("Updated formula in E3: " + cells["E3"].Formula);
            Console.WriteLine("Calculated value in E3: " + cells["E3"].Value);
        }
    }
}
