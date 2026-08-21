// Title: C# – Rename Named Ranges Containing “Old” to “New” and Recalculate Formulas with Aspose.Cells
// Description: This Aspose.Cells for .NET example loads a workbook, scans the NameCollection, replaces the substring “Old” with “New” in each defined name, updates the Name objects (automatically adjusting any formulas that reference them), triggers a full formula recalculation, and saves the modified file.
// Keywords: Aspose.Cells rename named ranges | C# rename Excel defined names | replace substring in NameCollection | recalculate formulas Aspose.Cells | batch rename Excel named ranges .NET | automatic formula update Aspose.Cells
// Common Searches: rename named ranges containing 'Old' Aspose.Cells C# | how to change defined names in Excel programmatically .NET | recalculate formulas after renaming named ranges Aspose.Cells | batch update Excel named ranges with Aspose.Cells | replace text in Excel named ranges using C#
// Developer Intent: Rename every named range that includes "Old" to "New" and refresh all dependent formulas.
// Use Cases: Standardize legacy named ranges across corporate workbooks | Migrate Excel templates to a new naming convention before deployment | Integrate named‑range renaming into automated build or CI pipelines | Prepare workbooks for localization by updating range identifiers
// AI Prompts: Write C# code using Aspose.Cells that changes every defined name containing "Temp" to "Current" and then forces a full formula recalculation. | Provide a unit test in C# that confirms formulas reference the new names after a named‑range rename with Aspose.Cells. | Explain how Aspose.Cells automatically updates formula references when the Name.Text property is modified.

using System;
using Aspose.Cells;

namespace RenameNamedRangesDemo
{
    // This Aspose.Cells for .NET example loads a workbook, scans the NameCollection, replaces the substring “Old” with “New” in each defined name, updates the Name objects (automatically adjusting any formulas that reference them), triggers a full formula recalculation, and saves the modified file.
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Get the collection of all defined names in the workbook
            NameCollection names = workbook.Worksheets.Names;

            // Iterate through each name and rename those containing "Old" to "New"
            for (int i = 0; i < names.Count; i++)
            {
                Name name = names[i];

                // Check if the name text includes the substring "Old"
                if (name.Text != null && name.Text.Contains("Old"))
                {
                    // Create the new name by replacing "Old" with "New"
                    string newName = name.Text.Replace("Old", "New");

                    // Assign the new name back to the Name object
                    // This updates the name and automatically adjusts formulas that reference it
                    name.Text = newName;
                }
            }

            // Recalculate all formulas so that any dependent calculations reflect the renamed ranges
            workbook.CalculateFormula();

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
