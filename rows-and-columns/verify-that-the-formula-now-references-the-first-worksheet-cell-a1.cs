// Title: How to verify that a cell formula points to the first worksheet's A1 cell using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# check if a formula references the first sheet A1 | detect A1 reference in Excel formula with Aspose.Cells .NET | C# code to validate that a cell formula uses Sheet1!A1 in Aspose.Cells | normalize Excel formula string and test for A1 reference using Aspose.Cells | verify Excel cell formula points to first worksheet cell A1 in C#
// Tags: Aspose.Cells formula A1 detection | C# Excel formula validation with Aspose.Cells | initial worksheet A1 reference detection | case‑insensitive formula parsing Aspose.Cells | qualified sheet name handling in Excel formulas

using Aspose.Cells;
using System;

// The example loads an Excel workbook, reads a specific cell's formula, normalizes the formula string for case‑insensitive comparison, and determines whether it directly references A1 or references A1 with the first worksheet's name, then outputs the verification result.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with the actual file path)
        Workbook workbook = new Workbook("input.xlsx"); // load rule

        // Get the first worksheet
        Worksheet firstSheet = workbook.Worksheets[0];

        // Cell that contains the formula to verify (adjust address as needed)
        Cell formulaCell = firstSheet.Cells["B1"]; // example cell

        // Retrieve the formula string
        string formula = formulaCell.Formula;

        // Flag indicating whether the formula references A1 on the first worksheet
        bool referencesA1 = false;

        if (!string.IsNullOrEmpty(formula))
        {
            // Normalize for case‑insensitive comparison
            string normalized = formula.Replace(" ", "").ToUpperInvariant();

            // Direct reference like "A1"
            bool directA1 = normalized.Contains("A1") && !normalized.Contains("!");

            // Qualified reference like "Sheet1!A1"
            string sheetName = firstSheet.Name.ToUpperInvariant();
            bool qualifiedA1 = normalized.Contains($"{sheetName}!A1");

            referencesA1 = directA1 || qualifiedA1;
        }

        // Output the verification result
        Console.WriteLine($"Formula in {formulaCell.Name} references first worksheet cell A1: {referencesA1}");
    }
}
