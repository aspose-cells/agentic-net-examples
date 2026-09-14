// Title: Replace Excel POWER() formulas with the ^ exponentiation operator using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, scans every cell, and rewrites any POWER(base, exponent) formula to use the caret (^) operator. | Create a regular‑expression routine in C# that transforms all POWER functions in worksheet formulas to the ^ syntax and saves the modified workbook using Aspose.Cells. | Write a method that iterates through a worksheet's cells, detects formulas containing POWER, substitutes them with base^exponent, and persists the changes via the Aspose.Cells API.
// Common Searches: how to convert POWER function to ^ operator in Aspose.Cells C# | replace Excel POWER formulas with caret using .NET library | regex replace POWER(base,exp) in workbook cells Aspose.Cells example
// Tags: POWER to caret conversion Aspose.Cells | formula regex replacement C# | Excel formula exponentiation operator Aspose.Cells | cell formula update .NET | Aspose.Cells workbook formula manipulation

using Aspose.Cells;
using System.Text.RegularExpressions;

// Load an existing workbook
Workbook workbook = new Workbook("input.xlsx");

// Access the first worksheet (adjust index as needed)
Worksheet worksheet = workbook.Worksheets[0];

// Iterate through all cells in the worksheet
foreach (Cell cell in worksheet.Cells)
{
    // Process only cells that contain a formula with the POWER function
    if (cell.IsFormula && cell.Formula.Contains("POWER"))
    {
        // Replace POWER(base, exponent) with base^exponent using a regular expression
        string updatedFormula = Regex.Replace(
            cell.Formula,
            @"POWER\(\s*([^,]+?)\s*,\s*([^\)]+?)\s*\)",
            "$1^$2",
            RegexOptions.IgnoreCase);

        // Assign the transformed formula back to the cell
        cell.Formula = updatedFormula;
    }
}

// Save the modified workbook
workbook.Save("output.xlsx");
