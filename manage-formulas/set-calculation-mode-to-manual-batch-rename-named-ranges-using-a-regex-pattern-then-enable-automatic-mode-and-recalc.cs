// Title: Batch Rename Named Ranges with Regex in Aspose.Cells .NET – Manual Calculation Mode
// Description: Demonstrates how to set Aspose.Cells calculation mode to Manual, rename all defined names using a regular‑expression pattern, switch back to Automatic mode, and force a full workbook recalculation in C#.
// Keywords: Aspose.Cells manual calculation mode | rename named ranges Aspose.Cells | regex batch rename defined names | C# Aspose.Cells calculate formula | disable auto calculation Aspose.Cells | Aspose.Cells .NET workbook rename | named range regex replace | Aspose.Cells calculation settings
// Common Searches: Aspose.Cells set calculation mode to manual | batch rename named ranges using regex Aspose.Cells | how to recalculate workbook after renaming names Aspose.Cells | disable automatic calculation Aspose.Cells .NET | rename all defined names Aspose.Cells C#
// Developer Intent: Temporarily turn off automatic formula calculation, rename every named range with a regex rule, then re‑enable automatic calculation and recalculate the workbook.
// Use Cases: Prevent unnecessary recalculations while performing bulk name changes. | Apply a new naming convention (e.g., replace a prefix) across many defined names. | Ensure all formulas reference the updated names by forcing a full recalculation after the rename.
// AI Prompts: Write C# code using Aspose.Cells that switches calculation to Manual, renames all defined names with a Regex replace, restores Automatic mode, and calls CalculateFormula. | Show an Aspose.Cells .NET example that batch updates named ranges without triggering intermediate recalculations. | Explain the steps to efficiently rename named ranges in a large workbook while temporarily disabling calculation and then refreshing all formulas.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Demonstrates how to set Aspose.Cells calculation mode to Manual, rename all defined names using a regular‑expression pattern, switch back to Automatic mode, and force a full workbook recalculation in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Sample data: create a worksheet and a named range
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Define a named range with a name that will be renamed later
            int nameIdx = workbook.Worksheets.Names.Add("OldRange1");
            Name oldName = workbook.Worksheets.Names[nameIdx];
            oldName.RefersTo = "=Sheet1!$A$1:$A$3";

            // -------------------------------------------------
            // 1. Set calculation mode to Manual
            // -------------------------------------------------
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // -------------------------------------------------
            // 2. Batch rename named ranges using a regex pattern
            // -------------------------------------------------
            // Example: replace the prefix "Old" with "New"
            string pattern = @"^Old";
            string replacement = "New";

            foreach (Name name in workbook.Worksheets.Names)
            {
                string newNameText = Regex.Replace(name.Text, pattern, replacement);
                if (!newNameText.Equals(name.Text, StringComparison.Ordinal))
                {
                    name.Text = newNameText; // rename the defined name
                }
            }

            // -------------------------------------------------
            // 3. Switch back to Automatic calculation mode
            // -------------------------------------------------
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

            // -------------------------------------------------
            // 4. Recalculate all formulas in the workbook
            // -------------------------------------------------
            // Use Workbook.CalculateFormula() to recalculate the entire workbook
            workbook.CalculateFormula();

            // -------------------------------------------------
            // Save the modified workbook
            // -------------------------------------------------
            workbook.Save("RenamedRanges.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
