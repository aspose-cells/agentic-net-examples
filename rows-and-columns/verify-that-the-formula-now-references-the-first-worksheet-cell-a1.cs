// Title: C# Aspose.Cells code to verify that a formula in a cell targets A1 on the first worksheet
// AI Prompts: Generate a C# method with Aspose.Cells that takes a Worksheet and a cell address and returns true if the formula in that cell targets A1 on the same sheet, handling both explicit sheet names and implicit references. | Write C# code that loads an Excel file, reads the formula from cell B1 of the first worksheet, and prints whether the formula targets A1 on that sheet using Aspose.Cells. | Create a reusable utility in C# using Aspose.Cells that parses any formula string and determines if it includes A1 as a target cell on a specified worksheet, covering external workbook scenarios.
// Common Searches: aspnet aspose.cells how to determine if a formula references cell A1 on the first sheet | c# check if Excel formula in B1 points to A1 using Aspose.Cells library | detect sheet name in formula with Aspose.Cells C# example | verify same‑sheet cell reference in Aspose.Cells C# code | aspose.cells parse formula to find A1 reference on specific worksheet
// Tags: aspose.cells verify formula reference | c# parse excel formula aspose.cells | check cell reference A1 worksheet | aspose.cells sheet name extraction from formula | excel formula validation c#

using Aspose.Cells;
using System;

// The example loads an Excel workbook, reads the formula in cell B1 of the first worksheet, normalizes it, and determines whether the formula targets cell A1 on that same worksheet, supporting both sheet‑qualified and unqualified references, then outputs the verification result.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with actual path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet ws = workbook.Worksheets[0];

        // Get the cell that contains the formula (example: B1)
        Cell formulaCell = ws.Cells["B1"];

        // Retrieve the formula string
        string formula = formulaCell.Formula;

        // Flag to indicate if the formula references A1 on the first worksheet
        bool referencesA1 = false;

        if (!string.IsNullOrEmpty(formula))
        {
            // Remove leading '=' and convert to upper case for uniform comparison
            string normalized = formula.TrimStart('=').ToUpperInvariant();

            // Check if the formula contains a reference to A1
            if (normalized.Contains("A1"))
            {
                // Determine if a sheet name is specified (e.g., 'Sheet1'!A1)
                int exclamIndex = normalized.IndexOf('!');
                if (exclamIndex > 0)
                {
                    // Extract the sheet name part (remove surrounding quotes if any)
                    string sheetPart = normalized.Substring(0, exclamIndex).Trim('\'');
                    // Compare with the name of the first worksheet
                    if (string.Equals(sheetPart, ws.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        referencesA1 = true;
                    }
                }
                else
                {
                    // No sheet name means the reference is to the same sheet (the first worksheet)
                    referencesA1 = true;
                }
            }
        }

        // Output the verification result
        Console.WriteLine($"Formula in B1: {formula}");
        Console.WriteLine($"References A1 on the first worksheet: {referencesA1}");
    }
}
