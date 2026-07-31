// Title: Lock formula cells and unlock input cells with Aspose.Cells for .NET (C#)
// Description: C# example that loads or creates an Excel workbook, scans the used range, marks cells containing formulas as locked and all other cells as unlocked, then protects the worksheet (optionally with a password) so only input cells remain editable.
// Keywords: Aspose.Cells | C# | Excel worksheet protection | lock formula cells | unlock data entry cells | IsLocked property | Cell.IsFormula | programmatic workbook security | protect sheet with password
// Common Searches: how to lock only formula cells using Aspose.Cells C# | unlock non‑formula cells while protecting a worksheet in .NET | Aspose.Cells set IsLocked based on formula presence | protect Excel sheet programmatically Aspose.Cells | C# lock formulas and allow data entry
// Developer Intent: Automatically protect a worksheet by locking every cell that contains a formula and leaving all other cells editable.
// Use Cases: Distribute a financial model where calculations are protected but users can enter raw data. | Create a data‑entry template that prevents accidental changes to computed results. | Generate reports for multiple collaborators while safeguarding formula integrity.
// AI Prompts: Generate C# code with Aspose.Cells that locks all formula cells, unlocks the rest, and applies worksheet protection. | Explain step‑by‑step how to iterate over a worksheet’s used range, detect Cell.IsFormula, set Style.IsLocked, and protect the sheet. | Show an example that loads an existing Excel file, locks formula cells, unlocks input cells, adds a password, and saves the protected workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsLockFormulaCells
{
    // C# example that loads or creates an Excel workbook, scans the used range, marks cells containing formulas as locked and all other cells as unlocked, then protects the worksheet (optionally with a password) so only input cells remain editable.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // replace with new Workbook("input.xlsx") to load

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // If the cell contains a formula, lock it; otherwise unlock it
                    bool lockCell = cell.IsFormula; // true for formula cells

                    // Get current style, modify the IsLocked property, and apply it back
                    Style style = cell.GetStyle();
                    style.IsLocked = lockCell;
                    cell.SetStyle(style);
                }
            }

            // Protect the worksheet so that the locking takes effect
            // You can set a password if desired
            sheet.Protect(ProtectionType.All, "myPassword", null);

            // Optional: allow users to select locked cells (helps navigation)
            sheet.Protection.AllowSelectingLockedCell = true;

            // Save the workbook
            workbook.Save("LockedFormulas.xlsx");
        }
    }
}
