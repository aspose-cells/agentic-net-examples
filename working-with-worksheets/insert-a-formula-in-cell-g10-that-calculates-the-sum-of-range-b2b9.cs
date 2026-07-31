// Title: C# AspNet Aspose.Cells: Insert a SUM formula in G10 to total B2:B9
// Description: Demonstrates how to create a workbook with Aspose.Cells, populate cells B2‑B9, assign the formula "=SUM(B2:B9)" to cell G10, recalculate the sheet, save the file, and read the computed result.
// Keywords: Aspose.Cells C# formula example | insert SUM formula Aspose.Cells | set cell formula .NET | calculate range total programmatically | G10 SUM B2:B9 Aspose | C# workbook automation | Aspose.Cells calculate formulas | save workbook after formula
// Common Searches: how to add a SUM formula to a cell using Aspose.Cells C# | Aspose.Cells set formula and recalculate workbook | C# code to sum a range in Excel with Aspose | write and evaluate Excel formula with Aspose.Cells .NET | save Excel file after inserting formula Aspose
// Developer Intent: Place a SUM expression in G10 that adds values from B2 through B9 and obtain the calculated total.
// Use Cases: Automatically compute totals for dynamically generated data tables. | Create financial statements where subtotals are derived via formulas before export. | Validate numeric results by reading the evaluated value after formula execution.
// AI Prompts: Generate C# Aspose.Cells code that writes "=SUM(B2:B9)" to G10, triggers calculation, and returns the result. | Show how to assign a formula to a cell, recalculate the workbook, and save it using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    // Demonstrates how to create a workbook with Aspose.Cells, populate cells B2‑B9, assign the formula "=SUM(B2:B9)" to cell G10, recalculate the sheet, save the file, and read the computed result.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Optionally put some sample values in B2:B9
            for (int i = 2; i <= 9; i++)
            {
                // Cells are zero‑based; column B is index 1
                worksheet.Cells[i - 1, 1].PutValue(i * 10); // Example values: 20,30,...,90
            }

            // Insert the SUM formula into cell G10 (column G = index 6, row 10 = index 9)
            worksheet.Cells[9, 6].Formula = "=SUM(B2:B9)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook to a file (lifecycle: save)
            workbook.Save("FormulaDemo.xlsx");

            // Output the result of the formula for verification
            Console.WriteLine("Result in G10: " + worksheet.Cells[9, 6].Value);
        }
    }
}
