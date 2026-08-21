// Title: Refactor Excel Formulas: Replace Hard‑Coded Numbers with Named Constants using Aspose.Cells for .NET
// Description: A C# utility that loads an Excel workbook, detects numeric literals in every formula (excluding cell references), creates unique named constants in the workbook’s name collection, substitutes the literals with those names, and saves the refactored file. Includes optional console output of the generated constants.
// Keywords: Aspose.Cells | C# | .NET | named constants | Excel formula refactoring | hard‑coded numbers | replace numeric literals | create named range | automate formula cleanup | Excel automation
// Common Searches: Aspose.Cells replace numeric literals with named constants | C# code to convert hard‑coded numbers in Excel formulas | how to create named ranges from constants in Aspose.Cells | refactor Excel formulas programmatically .NET | detect and rename constant values in worksheet formulas
// Developer Intent: Automatically convert hard‑coded numeric values in Excel formulas into reusable named constants.
// Use Cases: Standardize a tax rate (e.g., 0.07) across all worksheets by defining a single named constant. | Update a discount percentage (e.g., 15) in one place instead of editing each formula manually. | Prepare a financial model for multi‑currency or regional adjustments by extracting all literal numbers into named parameters.
// AI Prompts: Generate C# code with Aspose.Cells that scans every formula in a workbook and replaces each numeric literal with a unique named constant, then saves the result. | Explain how to extend the sample to also detect string literals and date constants and replace them with named ranges. | Provide a step‑by‑step verification checklist to ensure formulas recalculate correctly after substituting named constants.

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsFormulaRefactor
{
    // A C# utility that loads an Excel workbook, detects numeric literals in every formula (excluding cell references), creates unique named constants in the workbook’s name collection, substitutes the literals with those names, and saves the refactored file. Includes optional console output of the generated constants.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (replace with actual path)
            string inputPath = "input.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Dictionary to keep track of created named constants (constant value -> name)
            Dictionary<string, string> constantNames = new Dictionary<string, string>();

            // Regular expression to find numeric literals that are not part of cell references.
            // It matches numbers that are not preceded by a letter or '$' (to avoid A1, $B$2 etc.).
            Regex numberRegex = new Regex(@"(?<![A-Za-z\$])\b\d+(\.\d+)?\b", RegexOptions.Compiled);

            // Iterate over all used cells
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        string originalFormula = cell.Formula; // e.g. "=A1*5+10"
                        string updatedFormula = originalFormula;

                        // Find all numeric constants in the formula
                        MatchCollection matches = numberRegex.Matches(originalFormula);
                        foreach (Match match in matches)
                        {
                            string constantValue = match.Value; // e.g. "5" or "10"

                            // Skip if the constant is already part of a named range (e.g., =MyConst)
                            if (constantValue.StartsWith("="))
                                continue;

                            // Determine a unique name for this constant (e.g., Const_5, Const_10_1)
                            string baseName = "Const_" + constantValue.Replace(".", "_");
                            string constName = baseName;
                            int suffix = 1;
                            while (constantNames.ContainsValue(constName))
                            {
                                constName = baseName + "_" + suffix;
                                suffix++;
                            }

                            // If we haven't created a named constant for this value yet, add it
                            if (!constantNames.ContainsKey(constantValue))
                            {
                                // Add a new name to the workbook's name collection
                                int nameIndex = workbook.Worksheets.Names.Add(constName);
                                Name namedConst = workbook.Worksheets.Names[nameIndex];

                                // Define the constant as a literal (e.g., =5)
                                namedConst.RefersTo = "=" + constantValue;

                                // Store the mapping for future reuse
                                constantNames[constantValue] = constName;
                            }

                            // Replace the literal in the formula with the named constant
                            // Use word boundaries to avoid partial replacements
                            updatedFormula = Regex.Replace(updatedFormula,
                                                          $@"\b{Regex.Escape(constantValue)}\b",
                                                          constName);
                        }

                        // If the formula changed, apply the new formula using SetFormula (rule)
                        if (!updatedFormula.Equals(originalFormula, StringComparison.Ordinal))
                        {
                            // SetFormula(string formula, object value) – we let Aspose recalculate later, so pass null
                            cell.SetFormula(updatedFormula, null);
                        }
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save("output_refactored.xlsx");

            // Optional: display the created named constants
            Console.WriteLine("Created named constants:");
            foreach (var kvp in constantNames)
            {
                Console.WriteLine($"{kvp.Value} = {kvp.Key}");
            }
        }
    }
}
