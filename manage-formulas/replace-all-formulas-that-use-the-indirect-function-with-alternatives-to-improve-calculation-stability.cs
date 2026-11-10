// Title: Replace INDIRECT formulas with direct references in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to replace Excel INDIRECT formulas with static cell references using Aspose.Cells | how to remove INDIRECT function from formulas in a .xlsx file with Aspose.Cells library | Aspose.Cells replace INDIRECT("A1") pattern with direct reference in C# | update Excel formulas for stability by eliminating INDIRECT via Aspose.Cells | programmatically convert INDIRECT formulas to direct references in C# workbook
// Tags: Aspose.Cells formula transformation with regex | direct cell reference substitution C# | recalculate workbook after formula updates Aspose.Cells | Excel calculation stability improvement C# | eliminate INDIRECT usage Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The C# program loads an XLSX workbook with Aspose.Cells, scans every worksheet and cell for formulas containing the INDIRECT function, uses a case‑insensitive regular expression to replace INDIRECT("ref") patterns with the plain reference "ref", updates the cell formula when a change occurs, recalculates all formulas, and saves the modified workbook to a new file.
class ReplaceIndirectFormulas
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Regex to match INDIRECT functions with a quoted reference, e.g., INDIRECT("A1")
            Regex indirectRegex = new Regex(@"INDIRECT\(\s*""([^""]+)""\s*\)", RegexOptions.IgnoreCase);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    // Check if the cell contains a formula and the formula includes INDIRECT
                    if (cell.IsFormula && cell.Formula.Contains("INDIRECT", StringComparison.OrdinalIgnoreCase))
                    {
                        string originalFormula = cell.Formula;

                        // Replace each INDIRECT("ref") with the direct reference "ref"
                        string updatedFormula = indirectRegex.Replace(originalFormula, "$1");

                        // If a replacement was made, assign the new formula back to the cell
                        if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                        {
                            cell.Formula = updatedFormula;
                        }
                    }
                }
            }

            // Recalculate all formulas after modifications
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
