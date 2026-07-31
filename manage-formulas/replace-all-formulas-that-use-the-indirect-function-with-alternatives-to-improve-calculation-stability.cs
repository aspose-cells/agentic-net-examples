// Title: C# – Replace Excel INDIRECT Formulas with Direct References Using Aspose.Cells
// Description: Loads an Excel workbook, scans every worksheet and cell for formulas that contain the INDIRECT function, converts them to stable direct references (handling both quoted literals and cell‑based address values), recalculates the workbook, and saves the updated file.
// Keywords: Aspose.Cells | C# | replace INDIRECT formula | direct cell reference | Excel formula stability | convert INDIRECT to A1 | recalculate workbook | programmatic Excel modification
// Common Searches: how to remove INDIRECT function with Aspose.Cells C# | replace dynamic Excel references using Aspose.Cells | convert INDIRECT("A1") to =A1 programmatically | Aspose.Cells replace indirect formulas example | stable Excel formulas C# Aspose
// Developer Intent: Programmatically replace all INDIRECT formulas in an Excel workbook with equivalent direct references to improve calculation reliability.
// Use Cases: Transform string‑literal INDIRECT calls (e.g., =INDIRECT("B2")) into simple =B2 formulas. | Resolve INDIRECT formulas that point to a cell containing an address (e.g., =INDIRECT(C1)) and substitute the resolved address. | Recalculate the workbook after replacements to ensure dependent calculations reflect the new formulas.
// AI Prompts: Write C# code with Aspose.Cells that iterates through a workbook and replaces any INDIRECT formula with a direct reference, supporting quoted literals and cell‑based addresses. | Provide a robust TryReplaceIndirect method that validates extracted addresses, skips unsupported patterns, and logs cells that could not be converted.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsIndirectReplacement
{
    // Loads an Excel workbook, scans every worksheet and cell for formulas that contain the INDIRECT function, converts them to stable direct references (handling both quoted literals and cell‑based address values), recalculates the workbook, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (replace with actual path)
            string inputPath = "InputWorkbook.xlsx";
            // Path to the output workbook
            string outputPath = "OutputWorkbook.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate through all used cells
                foreach (Cell cell in cells)
                {
                    // Process only formula cells that contain the INDIRECT function
                    if (cell.IsFormula && cell.Formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string originalFormula = cell.Formula; // e.g. "=INDIRECT(\"A1\")" or "=INDIRECT(B1)"
                        string newFormula = TryReplaceIndirect(originalFormula, sheet);
                        if (!string.IsNullOrEmpty(newFormula) && newFormula != originalFormula)
                        {
                            // Set the new stable formula (lifecycle rule: set formula)
                            cell.Formula = newFormula;
                        }
                    }
                }
            }

            // Recalculate all formulas after replacement (feature rule: calculate)
            workbook.CalculateFormula();

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }

        static string TryReplaceIndirect(string formula, Worksheet sheet)
        {
            // Remove leading '=' if present for easier parsing
            string trimmed = formula.TrimStart('=').Trim();

            // Regex to capture the argument inside INDIRECT(...)
            Match match = Regex.Match(trimmed, @"INDIRECT\s*\(\s*(.+?)\s*\)", RegexOptions.IgnoreCase);
            if (!match.Success)
                return formula; // Not a recognizable INDIRECT usage

            string argument = match.Groups[1].Value;

            // Case 1: Argument is a quoted string literal e.g. "A1"
            if (argument.StartsWith("\"") && argument.EndsWith("\""))
            {
                string address = argument.Trim('\"');
                // Build direct reference formula
                return $"={address}";
            }

            // Case 2: Argument is a cell reference e.g. B1
            // Resolve the address stored in that cell
            Cell refCell = sheet.Cells[argument];
            if (refCell != null && refCell.Value != null)
            {
                string address = refCell.StringValue.Trim();
                // Basic validation: address should look like A1, B2, etc.
                if (Regex.IsMatch(address, @"^[A-Za-z]+\d+$"))
                {
                    return $"={address}";
                }
            }

            // If we cannot determine a stable address, return the original formula
            return formula;
        }
    }
}
