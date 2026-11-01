// Title: Verify English (US) function names and comma separators in Excel formulas with Aspose.Cells for .NET
// Common Searches: C# code to check Excel formulas for English function names using Aspose.Cells | How to detect semicolon argument separators in Excel formulas with Aspose.Cells .NET | Identify localized function names in a workbook programmatically with Aspose.Cells | Validate all formulas in an Excel file for US English syntax using Aspose.Cells | Iterate through cells and flag non‑English formulas in C# with Aspose.Cells
// Tags: Aspose.Cells English function name verification | semicolon separator detection in Excel formulas C# | lowercase function regex validation Aspose.Cells | worksheet cell iteration for formula checks .NET | save workbook after formula validation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The sample loads an Excel workbook, iterates over every cell in the first worksheet, and flags any formula that uses a semicolon or a function name beginning with a lowercase letter, indicating non‑English (US) syntax; it then reports the overall result and optionally saves the workbook.
class FormulaVerifier
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        bool allFormulasValid = true;

        // Iterate through all cells in the worksheet
        foreach (Cell cell in worksheet.Cells)
        {
            // Process only cells that contain a formula
            if (!string.IsNullOrEmpty(cell.Formula))
            {
                string formula = cell.Formula;

                // 1. Verify that commas are used as argument separators (no semicolons)
                if (formula.Contains(";"))
                {
                    Console.WriteLine($"Cell {cell.Name} uses semicolon as argument separator: {formula}");
                    allFormulasValid = false;
                }

                // 2. Verify that function names are English (US) – they are typically uppercase.
                // Flag any function name that starts with a lowercase letter (common in localized names).
                var lowerCaseFunctionMatches = Regex.Matches(formula, @"([a-z]+)\s*\(");
                if (lowerCaseFunctionMatches.Count > 0)
                {
                    Console.WriteLine($"Cell {cell.Name} may contain non‑English function names: {formula}");
                    allFormulasValid = false;
                }
            }
        }

        // Output overall verification result
        Console.WriteLine(allFormulasValid
            ? "All formulas use English (US) function names and commas."
            : "Some formulas need correction.");

        // Save the workbook (optional)
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
