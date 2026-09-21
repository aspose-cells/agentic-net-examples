// Title: How to lock formula cells and unlock input cells in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that iterates through each worksheet, locks cells that contain formulas, unlocks all other cells, and then protects the sheet before saving. | Show how to create an unlocked style, apply it to non‑formula cells, and enforce worksheet protection while keeping formula cells locked using Aspose.Cells.
// Common Searches: Aspose.Cells C# lock cells that contain formulas while allowing user entry in other cells | protect Excel worksheet but keep data entry cells editable with Aspose.Cells .NET | set IsLocked property for formula cells only using Aspose.Cells C# | unlock all non‑formula cells in an existing workbook and protect the sheet with Aspose.Cells
// Tags: Aspose.Cells formula cell locking | Aspose.Cells non‑formula cell unlocking | Aspose.Cells worksheet protection with styles | Aspose.Cells IsLocked property usage | Aspose.Cells iterate used range C#

using Aspose.Cells;

// The program loads an existing workbook, walks through each worksheet's used range, applies a locked style to cells that contain formulas, applies an unlocked style to all other cells, protects the worksheet, and saves the modified file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Prepare a style that marks cells as unlocked
            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false;

            // Loop through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    if (cell.IsFormula)
                    {
                        // Cells containing formulas should remain locked (default is locked)
                        // Ensure the style explicitly sets IsLocked = true
                        Style lockedStyle = cell.GetStyle();
                        lockedStyle.IsLocked = true;
                        cell.SetStyle(lockedStyle);
                    }
                    else
                    {
                        // Unlock all other cells for data entry
                        cell.SetStyle(unlockedStyle);
                    }
                }
            }

            // Protect the worksheet so that locked cells cannot be edited
            sheet.Protect(ProtectionType.All);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
