// Title: C# – Bulk replace CONCATENATE with CONCAT in Excel using Aspose.Cells
// Description: This Aspose.Cells for .NET example loads an Excel file, walks through every worksheet and used cell, detects formulas that call the legacy CONCATENATE function (case‑insensitive), swaps the function name to the modern CONCAT operator while keeping arguments intact, forces recalculation, and writes the updated workbook.
// Keywords: Aspose.Cells | C# | .NET | Excel formula conversion | CONCATENATE to CONCAT | bulk formula replace | programmatic Excel update | cell formula manipulation | modernize spreadsheets | Excel 2016+ | workbook automation
// Common Searches: Aspose.Cells replace CONCATENATE with CONCAT | C# code to change Excel CONCATENATE to CONCAT | bulk update Excel formulas using Aspose | convert legacy CONCATENATE function .NET | scan workbook and modify formulas Aspose.Cells
// Developer Intent: Convert all legacy CONCATENATE calls to the newer CONCAT syntax across a workbook.
// Use Cases: Upgrade legacy spreadsheets before sharing with newer Office versions | Automate batch processing of multiple workbooks to enforce consistent formula style | Integrate into CI pipelines that validate Excel files for modern functions | Reduce manual editing time for large datasets with concatenated strings
// AI Prompts: Write Aspose.Cells C# code that iterates over a workbook's used range and swaps every CONCATENATE call to CONCAT. | Create a robust helper that changes the function name in a formula string while preserving nested parentheses and arguments. | Suggest best‑practice error handling for loading, processing, and saving Excel files when performing bulk formula changes. | Explain how to trigger recalculation after modifying formulas with Aspose.Cells.

using System;
using Aspose.Cells;

// This Aspose.Cells for .NET example loads an Excel file, walks through every worksheet and used cell, detects formulas that call the legacy CONCATENATE function (case‑insensitive), swaps the function name to the modern CONCAT operator while keeping arguments intact, forces recalculation, and writes the updated workbook.
class ReplaceConcatOperator
{
    static void Main()
    {
        // Load the workbook (replace with actual path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Scan each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    string formula = cell.Formula;

                    // Check if the formula uses CONCATENATE (case‑insensitive)
                    if (!string.IsNullOrEmpty(formula) &&
                        formula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace CONCATENATE with CONCAT
                        string updatedFormula = ReplaceConcat(formula);

                        // Apply the new formula; value is set to null so Aspose recalculates it
                        cell.SetFormula(updatedFormula, null);
                    }
                }
            }
        }

        // Save the modified workbook (replace with desired output path)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }

    // Helper method to replace the function name while preserving the rest of the formula
    static string ReplaceConcat(string formula)
    {
        // Find the position of the function name (case‑insensitive)
        int index = formula.IndexOf("CONCATENATE", StringComparison.OrdinalIgnoreCase);
        if (index < 0) return formula;

        // Build the new formula string
        string before = formula.Substring(0, index);
        string after  = formula.Substring(index + "CONCATENATE".Length);
        return before + "CONCAT" + after;
    }
}
