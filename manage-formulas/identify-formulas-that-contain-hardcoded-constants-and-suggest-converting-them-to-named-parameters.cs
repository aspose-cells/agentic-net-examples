// Title: Aspose.Cells C# – Detect Hard‑Coded Numbers in Excel Formulas and Convert to Named Constants
// Description: A C# utility that loads an Excel workbook, scans every formula for numeric literals, creates a unique named constant for each distinct value, replaces the literals with those constants, recalculates the sheet, and saves the updated file. Ideal for improving formula maintainability and enabling parameter‑driven calculations.
// Keywords: Aspose.Cells formula parameterization | C# replace numeric literals with named constants | Excel hard‑coded numbers detection | named ranges for constants .NET | automate formula refactoring Aspose | Excel workbook constant extraction
// Common Searches: how to replace hard coded numbers in Excel formulas using Aspose.Cells | C# code to create named constants from numeric literals in workbook | Aspose.Cells find and substitute numeric values in formulas | convert Excel formula constants to named parameters .NET
// Developer Intent: Transform hard‑coded numeric values in worksheet formulas into reusable named constants to make the workbook easier to maintain and adapt.
// Use Cases: Bulk identification of all numeric literals across multiple worksheets. | Automatic generation of unique named constants (e.g., CONST_1) for each distinct number. | Updating formulas to reference the newly created constants instead of literal values. | Recalculating the workbook after changes and exporting the revised file.
// AI Prompts: Generate a C# method with Aspose.Cells that returns a list of unique numeric literals found in all formulas of a workbook. | Write code that adds a named constant for each numeric literal and rewrites the formulas to use those names, following the provided example. | Explain how to extend the regular expression to also match scientific notation such as 1.2E‑3 when detecting numbers in Excel formulas.

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsFormulaParameterizer
{
    // A C# utility that loads an Excel workbook, scans every formula for numeric literals, creates a unique named constant for each distinct value, replaces the literals with those constants, recalculates the sheet, and saves the updated file. Ideal for improving formula maintainability and enabling parameter‑driven calculations.
    class Program
    {
        // Regex to match numeric literals (integers or decimals) that are not part of cell references.
        // It looks for numbers that are preceded/followed by non-word characters or string boundaries.
        private static readonly Regex NumberRegex = new Regex(@"(?<![A-Za-z])(-?\d+(\.\d+)?)(?![A-Za-z0-9])",
                                                               RegexOptions.Compiled);

        static void Main(string[] args)
        {
            // Path to the source workbook (change as needed)
            string inputPath = "InputWorkbook.xlsx";
            // Path for the output workbook
            string outputPath = "OutputWorkbook_WithParameters.xlsx";

            // -------------------- Lifecycle: Create / Load --------------------
            // Load an existing workbook
            Workbook workbook = new Workbook(inputPath);
            // ------------------------------------------------------------------

            // Dictionary to keep track of created named constants
            Dictionary<string, string> constantNameMap = new Dictionary<string, string>();
            int constCounter = 1;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate through all used cells
                foreach (Cell cell in cells)
                {
                    // Process only cells that contain a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        string originalFormula = cell.Formula;
                        string updatedFormula = originalFormula;
                        bool formulaChanged = false;

                        // Find all numeric literals in the formula
                        foreach (Match match in NumberRegex.Matches(originalFormula))
                        {
                            string numberLiteral = match.Value;

                            // Skip if the number is part of a function name (e.g., "SUM1")
                            // Simple heuristic: ensure the character before the match is not a letter
                            // (already handled by the regex negative lookbehind).

                            // Get or create a named constant for this number
                            if (!constantNameMap.TryGetValue(numberLiteral, out string constName))
                            {
                                constName = $"CONST_{constCounter++}";
                                // Add the named constant to the workbook
                                int nameIndex = workbook.Worksheets.Names.Add(constName);
                                Name namedConst = workbook.Worksheets.Names[nameIndex];
                                // Named constants are defined with an equals sign, e.g., =5
                                namedConst.RefersTo = $"={numberLiteral}";
                                constantNameMap[numberLiteral] = constName;
                            }

                            // Replace the literal with the named constant in the formula
                            // Use word boundaries to avoid partial replacements
                            updatedFormula = Regex.Replace(updatedFormula,
                                                          $@"(?<![A-Za-z0-9_]){Regex.Escape(numberLiteral)}(?![A-Za-z0-9_])",
                                                          constName);
                            formulaChanged = true;
                        }

                        // If any replacement occurred, set the new formula back to the cell
                        if (formulaChanged && updatedFormula != originalFormula)
                        {
                            // Use SetFormula(string, object) overload; pass null for the pre‑calculated value.
                            cell.SetFormula(updatedFormula, null);
                        }
                    }
                }
            }

            // Optional: recalculate all formulas after modifications
            workbook.CalculateFormula();

            // -------------------- Lifecycle: Save --------------------
            workbook.Save(outputPath);
            // ---------------------------------------------------------

            Console.WriteLine($"Processing complete. Modified workbook saved to '{outputPath}'.");
        }
    }
}
