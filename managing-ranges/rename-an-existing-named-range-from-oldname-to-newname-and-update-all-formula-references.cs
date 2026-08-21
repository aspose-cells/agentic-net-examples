// Title: C# AspNet: Rename an Excel Named Range and Update All Formula References with Aspose.Cells
// Description: Loads an Excel file, finds the defined name "OldName", captures its RefersTo address, changes the name to "NewName", scans every worksheet to replace the old name in formulas, restores the original range reference, and saves the modified workbook.
// Keywords: Aspose.Cells rename named range C# | update formula references Aspose.Cells | change defined name programmatically | preserve RefersTo property | .NET Excel named range manipulation | bulk formula update Aspose.Cells
// Common Searches: how to rename a named range in Aspose.Cells .NET | replace old named range in all formulas using C# | preserve RefersTo while renaming Excel defined name | bulk update of named range references Aspose.Cells
// Developer Intent: Change the name of an existing Excel named range from "OldName" to "NewName" and automatically adjust every formula that references the old name.
// Use Cases: Standardize naming conventions after data model redesign without breaking dependent formulas. | Migrate legacy workbooks that use outdated named ranges to a new schema in an automated pipeline. | Synchronize workbook naming standards across multiple regional offices while keeping calculations intact.
// AI Prompts: Generate C# code with Aspose.Cells that renames a named range and rewrites all formula references in a workbook. | Explain how to keep the RefersTo address unchanged when renaming a defined name using Aspose.Cells. | Create a reusable Aspose.Cells method that accepts oldName and newName parameters, renames the range, and updates formulas across all worksheets.

using System;
using Aspose.Cells;

// Loads an Excel file, finds the defined name "OldName", captures its RefersTo address, changes the name to "NewName", scans every worksheet to replace the old name in formulas, restores the original range reference, and saves the modified workbook.
class RenameNamedRange
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Locate the named range with the old name
        Name oldName = null;
        foreach (Name name in workbook.Worksheets.Names)
        {
            if (name.Text == "OldName")
            {
                oldName = name;
                break;
            }
        }

        if (oldName != null)
        {
            // Preserve the original reference (e.g., "=Sheet1!$A$1:$A$10")
            string originalRefersTo = oldName.RefersTo;

            // Rename the defined name
            oldName.Text = "NewName";

            // Update all formulas that reference the old name
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula) && cell.Formula.Contains("OldName"))
                    {
                        cell.Formula = cell.Formula.Replace("OldName", "NewName");
                    }
                }
            }

            // Optionally, ensure the renamed name still points to the same range
            oldName.RefersTo = originalRefersTo;
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
