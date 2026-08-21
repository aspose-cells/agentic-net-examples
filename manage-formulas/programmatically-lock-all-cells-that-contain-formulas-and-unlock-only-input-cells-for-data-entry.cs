// Title: C# – Lock formula cells and unlock input cells with Aspose.Cells .NET
// Description: Shows how to create or load a workbook, detect cells that contain formulas, set their Locked style, unlock all other cells, protect the worksheet with a password, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | lock formula cells | unlock input cells | worksheet protection | StyleFlag | IsFormula | cell style locked attribute | Excel template security | protect sheet password
// Common Searches: lock only formula cells Aspose.Cells C# | unlock data entry cells while protecting worksheet Aspose.Cells | apply locked style based on formula detection Aspose.Cells | protect Excel sheet with password using Aspose.Cells | set cell IsLocked property in .NET Aspose.Cells
// Developer Intent: The developer wants to automatically lock every cell that contains a formula, unlock all other cells, and then protect the worksheet with a password.
// Use Cases: Create a spreadsheet template where end‑users can edit only designated input cells while all calculation cells stay protected. | Generate financial reports that lock formula cells to prevent accidental changes but allow users to enter new data in specific fields. | Distribute a workbook to collaborators with protected formulas and unlocked cells for data entry, secured by a worksheet password.
// AI Prompts: Provide C# code that iterates over the used range in Aspose.Cells, locks cells where IsFormula is true, unlocks the rest, and protects the worksheet with a password. | Explain how to use StyleFlag to apply only the Locked attribute when updating cell styles in Aspose.Cells for .NET. | Give a step‑by‑step guide to build a protected Excel template that allows editing only in input cells while keeping all formula cells locked.

using System;
using Aspose.Cells;

namespace AsposeCellsLockFormulaCells
{
    // Shows how to create or load a workbook, detect cells that contain formulas, set their Locked style, unlock all other cells, protect the worksheet with a password, and save the file using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data with formulas for demonstration
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].Formula = "=A1+A2";          // Formula cell
            cells["B2"].Formula = "=A1*2";           // Formula cell
            cells["C1"].PutValue("Input");           // Input cell
            cells["C2"].PutValue("Input");           // Input cell

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    Style style = cell.GetStyle();

                    // Lock cells that contain formulas, unlock others
                    if (cell.IsFormula)
                    {
                        style.IsLocked = true;   // Formula cells should be locked
                    }
                    else
                    {
                        style.IsLocked = false;  // Input cells should be unlocked
                    }

                    // Apply the locked/unlocked setting using a StyleFlag
                    StyleFlag flag = new StyleFlag();
                    flag.Locked = true;   // Apply the Locked attribute only
                    cell.SetStyle(style, flag);
                }
            }

            // Protect the worksheet so that the locking takes effect
            // Use a password and protect all aspects of the sheet
            sheet.Protect(ProtectionType.All, "securePassword", null);

            // Save the workbook
            workbook.Save("LockedFormulaCells.xlsx");
        }
    }
}
