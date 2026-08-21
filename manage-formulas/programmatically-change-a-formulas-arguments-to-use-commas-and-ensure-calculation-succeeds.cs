// Title: Replace ';' with ',' in Excel formulas using Aspose.Cells for .NET and recalculate the workbook
// Description: Loads a workbook, scans every cell for formulas that use the locale‑specific semicolon argument separator, swaps it for a comma, triggers a full calculation, and saves the corrected file. Ideal for normalizing formulas across regions.
// Keywords: Aspose.Cells formula separator | C# replace semicolon with comma in Excel formula | locale specific formula delimiter | batch update Excel formulas .NET | recalculate workbook Aspose.Cells | Excel formula delimiter conversion | global Excel compatibility
// Common Searches: Aspose.Cells change formula delimiter from ; to , | C# fix Excel formula separators programmatically | How to recalculate workbook after editing formulas Aspose | Convert locale specific Excel formulas using Aspose.Cells | Batch replace semicolons in Excel formulas .NET
// Developer Intent: Programmatically convert semicolon argument separators to commas in all formulas and force workbook recalculation.
// Use Cases: Standardize formulas after importing spreadsheets from European locales that use ';' as the argument separator. | Prepare a multi‑sheet report for distribution to users with different regional settings. | Automate formula cleanup in a CI pipeline that generates Excel files for downstream analytics.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through every cell, replaces ';' with ',' in formulas, recalculates, and saves the workbook. | Explain how Aspose.Cells handles locale‑dependent formula delimiters and the safest way to modify them. | Provide a step‑by‑step tutorial for bulk updating Excel formula separators and triggering calculation using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaCommaFix
{
    // Loads a workbook, scans every cell for formulas that use the locale‑specific semicolon argument separator, swaps it for a comma, triggers a full calculation, and saves the corrected file. Ideal for normalizing formulas across regions.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (you can iterate all worksheets if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Iterate through all used cells and replace semicolon argument separators with commas
            foreach (Cell cell in cells)
            {
                // Only process cells that already contain a formula
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Replace locale‑specific ';' with the standard ',' delimiter
                    string correctedFormula = cell.Formula.Replace(';', ',');

                    // Update the cell formula only if a change was made
                    if (correctedFormula != cell.Formula)
                    {
                        cell.Formula = correctedFormula;
                    }
                }
            }

            // Recalculate all formulas to ensure the workbook reflects the updated formulas
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
