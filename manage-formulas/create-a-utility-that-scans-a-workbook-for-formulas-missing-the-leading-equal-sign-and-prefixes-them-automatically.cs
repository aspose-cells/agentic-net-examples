// Title: C# utility to auto‑prepend missing ‘=’ to Excel formulas using Aspose.Cells
// Description: A concise C# example that loads an Excel workbook with Aspose.Cells, scans every worksheet and used cell, adds a leading equal sign to formulas that lack it, recalculates the workbook, and saves the corrected file.
// Keywords: Aspose.Cells | C# | add missing equal sign | fix Excel formulas | batch process workbooks | recalculate formulas | Excel automation | formula validation
// Common Searches: add leading = to Excel formulas C# Aspose.Cells | scan workbook for formulas without equal sign | auto‑fix missing = in Excel cells using Aspose | recalculate workbook after fixing formulas | batch correct Excel formulas programmatically
// Developer Intent: Identify formula cells missing the leading '=', prepend '=', recalculate, and save the workbook.
// Use Cases: Clean user‑generated spreadsheets where formulas were entered without ‘=’ before analysis. | Prepare imported Excel files for reliable calculations in automated reporting pipelines. | Batch‑process multiple workbooks to ensure syntactically correct formulas prior to distribution.
// AI Prompts: Write C# code with Aspose.Cells that scans all cells and adds a leading ‘=’ to any formula missing it. | Show how to log the addresses of cells corrected while fixing formulas using Aspose.Cells. | Explain how to limit the utility to specific worksheets or a defined cell range.

using System;
using Aspose.Cells;

namespace FormulaFixUtility
{
    // A concise C# example that loads an Excel workbook with Aspose.Cells, scans every worksheet and used cell, adds a leading equal sign to formulas that lack it, recalculates the workbook, and saves the corrected file.
    public static class FormulaFixer
    {
        /// <summary>
        /// Loads a workbook, fixes malformed formulas, recalculates, and saves.
        /// </summary>
        /// <param name="inputPath">Path to the source Excel file.</param>
        /// <param name="outputPath">Path where the corrected file will be saved.</param>
        public static void Process(string inputPath, string outputPath)
        {
            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all used cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Process only cells that Aspose identifies as formulas
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;

                        // If the formula does not start with '=', prefix it
                        if (!string.IsNullOrEmpty(formula) && !formula.StartsWith("="))
                        {
                            cell.Formula = "=" + formula;
                        }
                    }
                }
            }

            // Recalculate all formulas after fixing (feature rule: CalculateFormula)
            workbook.CalculateFormula();

            // Save the corrected workbook (lifecycle rule: save)
            workbook.Save(outputPath);
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string sourceFile = "input.xlsx";
            string correctedFile = "output_fixed.xlsx";

            FormulaFixer.Process(sourceFile, correctedFile);

            Console.WriteLine("Formula scan and fix completed.");
        }
    }
}
