// Title: C# Aspose.Cells utility to detect and prepend missing ‘=’ in Excel formulas
// Description: A C# example that loads an Excel workbook with Aspose.Cells, scans every used cell, identifies formulas that do not start with the equal sign, prefixes the missing '=', recalculates all formulas, and saves the corrected file.
// Keywords: Aspose.Cells formula fix | C# add missing equal sign | prepend = to Excel formulas | bulk formula correction | scan workbook cells Aspose | Excel formula validation C# | recalculate formulas after edit | FormulaFixer utility | Excel automation missing = | Aspose.Cells FormulaFixer
// Common Searches: how to add a leading = to Excel formulas using Aspose.Cells | C# code to find and fix formulas without = in a workbook | Aspose.Cells scan cells for malformed formulas | bulk correct missing equal signs in Excel with C# | recalculate workbook after fixing formulas Aspose | detect plain‑text formulas in Excel files
// Developer Intent: Automatically locate cells whose Formula property lacks a leading '=', prepend the '=', recalculate the workbook, and save the updated file.
// Use Cases: Clean up spreadsheets imported from systems that store formulas as plain text. | Prepare workbooks for accurate calculation before running analytics or reporting. | Integrate into migration pipelines to ensure formula syntax consistency across multiple files.
// AI Prompts: Generate C# code with Aspose.Cells that scans all worksheets, adds a missing '=' to any formula, recalculates, and saves the workbook. | Show how to log the addresses of cells that were corrected when fixing formulas using Aspose.Cells. | Explain how to modify the utility to target a specific worksheet or a defined cell range.

using System;
using Aspose.Cells;

namespace FormulaUtility
{
    // A C# example that loads an Excel workbook with Aspose.Cells, scans every used cell, identifies formulas that do not start with the equal sign, prefixes the missing '=', recalculates all formulas, and saves the corrected file.
    public static class FormulaFixer
    {
        public static void FixFormulas(string inputPath, string outputPath)
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Retrieve the formula string; empty if the cell has no formula
                        string formula = cell.Formula;

                        // If a formula exists but does not start with '=', fix it
                        if (!string.IsNullOrEmpty(formula) && !formula.StartsWith("="))
                        {
                            // Prefix the missing equal sign
                            cell.Formula = "=" + formula;
                        }
                    }
                }
            }

            // Recalculate all formulas after fixing (optional but recommended)
            workbook.CalculateFormula();

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }

        // Example usage
        public static void Main()
        {
            string sourceFile = "input.xlsx";
            string destinationFile = "output_fixed.xlsx";

            FixFormulas(sourceFile, destinationFile);

            Console.WriteLine($"Formulas fixed and workbook saved to '{destinationFile}'.");
        }
    }
}
