// Title: Rename Named Ranges Containing 'Old' to 'New' and Refresh Formulas with Aspose.Cells for .NET
// Description: Loads an Excel workbook, renames every defined name that includes "Old" to "New", updates all formulas that reference the old names, recalculates the workbook, and saves the result using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | rename named ranges | replace substring in defined names | update formulas | recalculate workbook | Excel automation | .NET Excel library | batch rename ranges | formula refresh
// Common Searches: Aspose.Cells rename named range Old to New | C# update formulas after renaming defined name | replace part of a named range in Excel using Aspose | batch rename Excel named ranges with Aspose.Cells | recalculate workbook after name change Aspose
// Developer Intent: Rename all named ranges that contain "Old" to "New", adjust any formulas that reference them, and recalculate the workbook.
// Use Cases: Standardize legacy named ranges in financial models before a migration to a new naming convention. | Automate the update of named ranges after a corporate rebranding to keep all formulas valid. | Apply a consistent prefix to named ranges across multiple workbooks prior to distribution.
// AI Prompts: Generate C# code with Aspose.Cells that renames every defined name containing "Temp" to "Final" and updates all dependent formulas. | Create a reusable method that scans a workbook for named ranges matching a pattern, renames them, updates referencing formulas, and triggers a full calculation.

using System;
using Aspose.Cells;

namespace RenameNamedRangesDemo
{
    // Loads an Excel workbook, renames every defined name that includes "Old" to "New", updates all formulas that reference the old names, recalculates the workbook, and saves the result using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all defined names in the workbook
            foreach (Name name in workbook.Worksheets.Names)
            {
                // Check if the name contains the substring "Old"
                if (name.Text != null && name.Text.Contains("Old"))
                {
                    string oldName = name.Text;
                    string newName = oldName.Replace("Old", "New");

                    // Rename the defined name
                    name.Text = newName;

                    // Update all formulas that reference the old name
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        Cells cells = ws.Cells;
                        foreach (Cell cell in cells)
                        {
                            if (cell.IsFormula && cell.Formula != null && cell.Formula.Contains(oldName))
                            {
                                string updatedFormula = cell.Formula.Replace(oldName, newName);
                                cell.Formula = updatedFormula;
                            }
                        }
                    }
                }
            }

            // Recalculate all formulas to reflect the renamed ranges
            workbook.CalculateFormula();

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
