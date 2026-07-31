// Title: Lock Formula Cells and Protect Worksheet with Aspose.Cells (C#)
// Description: Creates a workbook, inserts values and formulas, locks every cell that contains a formula, unlocks non‑formula cells, applies full worksheet protection, and saves the file as a protected Excel document.
// Keywords: Aspose.Cells lock formula cells | protect worksheet C# | lock cells based on IsFormula | prevent editing formulas Aspose | Excel cell protection Aspose.Cells | C# Aspose.Cells worksheet protection
// Common Searches: How to lock only formula cells using Aspose.Cells for .NET | Protect a worksheet while keeping input cells editable in Aspose.Cells | Lock cells that contain formulas in an Excel file with Aspose.Cells | Aspose.Cells lock cells with formulas example
// Developer Intent: Automatically lock all formula cells and protect the sheet so end users cannot modify calculated values.
// Use Cases: Generate a calculation workbook, then lock derived cells while leaving input cells editable. | Create a reusable template where users can enter data but cannot alter the underlying formulas. | Distribute a final report with all formulas protected to ensure result integrity.
// AI Prompts: Show C# code using Aspose.Cells to lock formula cells, unlock other cells, and protect the worksheet. | Give an example that iterates through used cells, sets IsLocked based on IsFormula, and saves a protected workbook. | Explain how to apply worksheet protection after locking specific cells with Aspose.Cells and configure protection options.

using System;
using Aspose.Cells;

// Creates a workbook, inserts values and formulas, locks every cell that contains a formula, unlocks non‑formula cells, applies full worksheet protection, and saves the file as a protected Excel document.
class LockFormulaCells
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data and formulas
        cells["A1"].PutValue(10);
        cells["A2"].PutValue(20);
        cells["B1"].Formula = "=A1+A2";   // formula cell
        cells["B2"].Formula = "=A1*2";    // formula cell
        cells["C1"].PutValue("Sample");  // non‑formula cell

        // Loop through all used cells
        foreach (Cell cell in cells)
        {
            // If the cell contains a formula, lock it
            if (cell.IsFormula)
            {
                Style style = cell.GetStyle();
                style.IsLocked = true;          // lock formula cell
                cell.SetStyle(style);
            }
            else
            {
                // Optional: unlock cells without formulas so they remain editable
                Style style = cell.GetStyle();
                style.IsLocked = false;
                cell.SetStyle(style);
            }
        }

        // Protect the worksheet (locking takes effect only when protected)
        worksheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("LockedFormulas.xlsx");
    }
}
