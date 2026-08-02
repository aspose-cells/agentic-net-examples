// Title: Aspose.Cells .NET: Set Manual Calculation, Batch Rename Named Ranges with Regex, Restore Automatic Mode & Recalculate
// Description: Demonstrates how to switch a workbook to manual calculation, rename all defined names using a regular‑expression pattern (e.g., add a "New_" prefix and replace spaces with underscores), then revert to automatic mode, trigger a full formula recalculation, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | manual calculation mode | automatic calculation mode | rename named ranges | batch rename defined names | regex rename | regular expression | Workbook.CalculateFormula | performance optimization | named range prefix | spaces to underscores
// Common Searches: Aspose.Cells set calculation mode manual | batch rename named ranges regex C# | how to rename defined names without recalculating | restore automatic calculation after renaming names Aspose.Cells | force formula recalculation after changing named ranges
// Developer Intent: Temporarily disable automatic formula calculation, rename all workbook named ranges using a regex pattern, then re‑enable automatic calculation and recalculate formulas.
// Use Cases: Improve performance by turning off automatic calculation while performing bulk renaming of named ranges. | Standardize naming conventions across a workbook with regex‑based transformations (e.g., prefix and replace spaces). | Ensure formulas reference the new names correctly by switching back to automatic mode and invoking CalculateFormula.
// AI Prompts: Generate C# code with Aspose.Cells that sets the calculation mode to Manual, renames every named range by adding "New_" and converting spaces to underscores using Regex, then switches back to Automatic and recalculates all formulas. | Explain how to batch rename defined names in Aspose.Cells without triggering intermediate formula recalculations, and why manual mode improves speed. | Show how to validate that formulas correctly reference the renamed named ranges after calling Workbook.CalculateFormula.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsRenameNamedRanges
{
    // Demonstrates how to switch a workbook to manual calculation, rename all defined names using a regular‑expression pattern (e.g., add a "New_" prefix and replace spaces with underscores), then revert to automatic mode, trigger a full formula recalculation, and save the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook (or load an existing one)
                // -------------------------------------------------
                Workbook workbook = new Workbook();

                // Add sample data and some named ranges for demonstration
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Create a few named ranges with **valid** names (no spaces)
                int idx1 = workbook.Worksheets.Names.Add("Total_Sales");
                workbook.Worksheets.Names[idx1].RefersTo = "=Sheet1!$A$1:$A$3";

                int idx2 = workbook.Worksheets.Names.Add("Average_Score");
                workbook.Worksheets.Names[idx2].RefersTo = "=Sheet1!$A$1:$A$3";

                // -------------------------------------------------
                // 2. Switch calculation mode to Manual
                // -------------------------------------------------
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                // -------------------------------------------------
                // 3. Batch rename named ranges using a regex pattern
                //    Example: replace spaces with underscores and prepend "New_"
                // -------------------------------------------------
                string pattern = @"\s+";
                string replacement = "_";

                foreach (Name name in workbook.Worksheets.Names)
                {
                    // Apply regex replacement to the existing name text
                    string newName = "New_" + Regex.Replace(name.Text, pattern, replacement);
                    name.Text = newName; // rename the defined name
                }

                // -------------------------------------------------
                // 4. Switch calculation mode back to Automatic
                // -------------------------------------------------
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

                // -------------------------------------------------
                // 5. Recalculate all formulas in the workbook
                // -------------------------------------------------
                workbook.CalculateFormula();

                // -------------------------------------------------
                // 6. Save the workbook
                // -------------------------------------------------
                string outputPath = "RenamedNamedRanges.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
