// Title: C# – Print cell E3 formula and its SUM result using Aspose.Cells for .NET
// Description: The sample creates a workbook, writes 10 to A1 and 20 to A2, sets the formula "=SUM(A1:A2)" in cell E3, triggers calculation, and outputs both the formula text and the evaluated value (30) to the console.
// Keywords: Aspose.Cells | C# | .NET | cell formula retrieval | calculate workbook formulas | console output | SUM function | worksheet.Cells | Workbook.CalculateFormula | E3 cell
// Common Searches: Aspose.Cells get formula text after calculation | C# print formula and result from Excel cell | Aspose.Cells display evaluated value in console | How to read cell formula with Aspose.Cells .NET | Sum formula example Aspose.Cells C#
// Developer Intent: Show the formula stored in a specific cell and its computed value in a console application.
// Use Cases: Debugging spreadsheet logic by logging formulas and results. | Generating quick console reports of key calculation cells. | Automated unit tests that verify formula outcomes.
// AI Prompts: Generate code to iterate over a range and print each cell's formula with its calculated value using Aspose.Cells. | Provide a C# example that writes a cell's formula and result to a text file instead of the console. | Explain how to format the printed result as currency when displaying a formula's value.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    // The sample creates a workbook, writes 10 to A1 and 20 to A2, sets the formula "=SUM(A1:A2)" in cell E3, triggers calculation, and outputs both the formula text and the evaluated value (30) to the console.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare some data that the formula will use
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);

            // Set a formula in cell E3
            worksheet.Cells["E3"].Formula = "=SUM(A1:A2)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the formula text and its calculated value
            Console.WriteLine("Formula in E3: " + worksheet.Cells["E3"].Formula);
            Console.WriteLine("Calculated value in E3: " + worksheet.Cells["E3"].Value);
        }
    }
}
