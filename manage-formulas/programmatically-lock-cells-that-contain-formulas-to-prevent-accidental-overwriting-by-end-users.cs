// Title: Lock Formula Cells and Protect Worksheet Using AspNet Cells for .NET (C#)
// Description: Shows how to create a workbook, insert values and formulas, lock only the formula cells, unlock other cells, apply worksheet protection (with optional password), and save the file with Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | lock formula cells | worksheet protection | Excel protection .NET | prevent formula editing | IsLocked property | Cell.IsFormula | protect worksheet password | Aspose.Cells example
// Common Searches: Aspose.Cells lock formula cells C# | protect worksheet while allowing data entry Aspose.Cells | how to lock only formula cells in Excel using .NET | set IsLocked for formula cells Aspose.Cells | Excel file protection Aspose.Cells C#
// Developer Intent: Programmatically lock cells that contain formulas and protect the worksheet so end users cannot overwrite calculated results.
// Use Cases: Financial models where calculation results must stay immutable while input cells remain editable. | Spreadsheet templates distributed to clients that require formula protection but allow data entry. | Automated report generation that secures all derived values before sharing the file.
// AI Prompts: Generate C# code with Aspose.Cells that locks only cells containing formulas, unlocks other cells, applies worksheet protection with a password, and saves the workbook. | Explain how to iterate over a worksheet's used range, detect Cell.IsFormula, set Style.IsLocked accordingly, and protect the sheet in Aspose.Cells. | Provide step‑by‑step guidance for protecting an Aspose.Cells worksheet after locking formula cells, including optional password usage.

using System;
using Aspose.Cells;

namespace AsposeCellsLockFormulaCells
{
    // Shows how to create a workbook, insert values and formulas, lock only the formula cells, unlock other cells, apply worksheet protection (with optional password), and save the file with Aspose.Cells for C#.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Add formulas in column B
            cells["B1"].Formula = "=A1*2";
            cells["B2"].Formula = "=A2*2";
            cells["B3"].Formula = "=A3*2";

            // Add a formula in C1 for demonstration
            cells["C1"].Formula = "=SUM(A1:A3)";

            // Iterate through used cells and lock only those that contain formulas
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    Style style = cell.GetStyle();

                    if (cell.IsFormula)
                    {
                        // Lock cells that have formulas
                        style.IsLocked = true;
                    }
                    else
                    {
                        // Unlock cells without formulas (optional, makes them editable)
                        style.IsLocked = false;
                    }

                    cell.SetStyle(style);
                }
            }

            // Protect the worksheet so that locked cells cannot be edited
            // Using a password for demonstration; you can omit the password if not needed
            sheet.Protect(ProtectionType.All, "securePwd", null);

            // Save the workbook
            workbook.Save("LockedFormulaCells.xlsx");
        }
    }
}
