// Title: Print the SUM(A1:A2) formula and its calculated value from cell E3 with Aspose.Cells in a C# console app
// AI Prompts: Set the formula '=SUM(A1:A2)' in worksheet cell E3, invoke workbook.CalculateFormula(), and output both worksheet.Cells["E3"].Formula and worksheet.Cells["E3"].Value to the console. | Assign a SUM range formula to a specific cell, recalculate all formulas in the workbook, and display the formula string alongside its evaluated numeric result using Aspose.Cells for .NET. | Create a new workbook, populate A1 and A2, apply a SUM formula to E3, trigger calculation, and print the formula text and resulting value in a .NET console program.
// Common Searches: aspocells c# console display cell formula after calculation | how to retrieve evaluated value of a formula cell using Aspose.Cells .NET | example of setting SUM formula in Excel worksheet with Aspose.Cells and printing result | Aspose.Cells calculate and print cell E3 formula and value in console application
// Tags: Aspose.Cells calculate workbook formulas | C# set SUM formula in worksheet cell | Aspose.Cells output cell formula to console | retrieve evaluated cell value Aspose.Cells | console application display Excel formula result

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    // Demonstrates creating a workbook, inserting values into A1 and A2, assigning a SUM(A1:A2) formula to cell E3, calculating all formulas, and writing both the formula string and its evaluated value to the console using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells that will be referenced by the formula
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);

            // Set the formula in cell E3
            worksheet.Cells["E3"].Formula = "=SUM(A1:A2)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the formula text and its calculated value
            Console.WriteLine("Formula in E3: " + worksheet.Cells["E3"].Formula);
            Console.WriteLine("Calculated value in E3: " + worksheet.Cells["E3"].Value);
        }
    }
}
