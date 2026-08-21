// Title: Replace named range references (OldRange → NewRange) in all formulas with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, scans every worksheet and cell, detects formulas that contain the named range "OldRange", swaps it for "NewRange", writes the updated formula back, and saves the modified file.
// Keywords: Aspose.Cells replace named range | C# update formula references | change OldRange to NewRange | bulk formula edit Aspose.Cells | .NET Excel range substitution
// Common Searches: Aspose.Cells rename named range in formulas | C# replace OldRange with NewRange in Excel workbook | how to update all formula references using Aspose.Cells | bulk edit named range across worksheets .NET
// Developer Intent: Replace every occurrence of the named range "OldRange" with "NewRange" in formulas throughout an Excel workbook using Aspose.Cells.
// Use Cases: Refresh legacy reports after a named range is renamed. | Redirect formulas to a new data block without manual edits. | Automate preprocessing of workbooks before distribution or publishing.
// AI Prompts: Write C# code with Aspose.Cells that swaps a specific named range in all formulas across a workbook. | Suggest a performance‑optimized method to update formula references without scanning every cell. | Explain how to validate formulas after a named‑range substitution using Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel workbook, scans every worksheet and cell, detects formulas that contain the named range "OldRange", swaps it for "NewRange", writes the updated formula back, and saves the modified file.
class ReplaceOldRange
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through each cell that contains data
            foreach (Cell cell in sheet.Cells)
            {
                // Check if the cell contains a formula referencing "OldRange"
                if (cell.IsFormula && cell.Formula != null && cell.Formula.Contains("OldRange"))
                {
                    // Replace the reference and assign the updated formula back to the cell
                    string updatedFormula = cell.Formula.Replace("OldRange", "NewRange");
                    cell.Formula = updatedFormula;
                }
            }
        }

        // Save the workbook with the updated formulas
        workbook.Save("output.xlsx");
    }
}
