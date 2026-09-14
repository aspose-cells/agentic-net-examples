// Title: Verify Excel formulas use US English function names and commas as separators with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that scans every worksheet, identifies formula cells, and lists any formulas that contain semicolons or function names not in a US‑English list. | Create a reusable method that accepts a Workbook object and returns a collection of Cell objects whose formulas violate US English naming or comma‑separator rules, without altering the workbook. | Generate a console‑application script that logs each invalid formula with its cell address, sheet name, and the specific issue (semicolon separator or localized function) using Aspose.Cells in .NET.
// Common Searches: aspocells c# check formula for semicolon separator | detect localized Excel function names in workbook using Aspose.Cells | validate that Excel formulas use US English functions in .NET | programmatically ensure commas are used as argument separators in Excel formulas with Aspose.Cells | list cells with invalid formulas across all sheets using Aspose.Cells C#
// Tags: Aspose.Cells formula validation US English | detect semicolon separators in Excel formulas C# | verify English function names in workbook formulas | scan all worksheets for invalid formulas Aspose.Cells | report cells with non‑US Excel functions .NET

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook, iterates through every cell in each worksheet, checks each formula for semicolon argument separators and for function names outside a predefined US English list, outputs validation results with cell address and issue, and saves the workbook unchanged.
class FormulaVerifier
{
    // List of common English (US) function names.
    private static readonly HashSet<string> EnglishFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "SUM","AVERAGE","COUNT","MAX","MIN","IF","AND","OR","NOT","VLOOKUP","HLOOKUP",
        "INDEX","MATCH","LEFT","RIGHT","MID","LEN","TRIM","CONCATENATE","ROUND",
        "ROUNDUP","ROUNDDOWN","ABS","POWER","SQRT","DATE","TIME","NOW","TODAY",
        "YEAR","MONTH","DAY","TEXT","VALUE","ISNUMBER","ISTEXT","ISBLANK"
        // Add more functions as needed.
    };

    static void Main()
    {
        // Load an existing workbook (replace with your file path).
        Workbook workbook = new Workbook("Input.xlsx");

        // Iterate through all worksheets.
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the used range to limit iteration.
            var cells = sheet.Cells;
            var maxRow = cells.MaxDataRow;
            var maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula; // e.g. "=SUM(A1:A5, B1:B5)"
                        bool isValid = true;
                        List<string> issues = new List<string>();

                        // 1. Verify commas are used as argument separators (no semicolons).
                        if (formula.Contains(";"))
                        {
                            isValid = false;
                            issues.Add("Uses semicolon ';' as argument separator.");
                        }

                        // 2. Verify that function names are English (US).
                        // Extract tokens that appear before '('.
                        int index = 0;
                        while ((index = formula.IndexOf('(', index)) != -1)
                        {
                            // Find the start of the function name.
                            int start = index - 1;
                            while (start >= 0 && (char.IsLetter(formula[start]) || formula[start] == '_'))
                                start--;
                            start++; // Move to first letter.

                            string funcName = formula.Substring(start, index - start);
                            if (!EnglishFunctions.Contains(funcName))
                            {
                                isValid = false;
                                issues.Add($"Non‑English function name detected: '{funcName}'.");
                            }
                            index++; // Move past '(' for next search.
                        }

                        // Output verification result.
                        Console.WriteLine($"Cell {cell.Name} (Sheet: {sheet.Name})");
                        Console.WriteLine($"  Formula: {formula}");
                        Console.WriteLine($"  Valid: {isValid}");
                        if (!isValid)
                        {
                            foreach (string issue in issues)
                                Console.WriteLine($"  Issue: {issue}");
                        }
                        Console.WriteLine();
                    }
                }
            }
        }

        // Optionally, save the workbook after verification (no changes made here).
        workbook.Save("VerifiedOutput.xlsx");
    }
}
