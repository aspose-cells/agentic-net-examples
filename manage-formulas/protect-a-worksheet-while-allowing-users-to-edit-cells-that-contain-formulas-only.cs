// Title: C# – Protect an Aspose.Cells worksheet while allowing only formula cells to be edited
// Description: Creates a workbook, unlocks cells that contain formulas, locks all other cells, applies worksheet protection with a password, and saves the file so only formula cells remain editable using Aspose.Cells for .NET.
// Keywords: Aspose.Cells protect worksheet C# | unlock formula cells Aspose.Cells | lock non‑formula cells .NET | worksheet protection password Aspose.Cells | IsFormula style.IsLocked | Excel sheet security C#
// Common Searches: protect worksheet but allow editing of formula cells Aspose.Cells | unlock only formula cells in Excel using Aspose.Cells C# | C# code to lock all cells except those with formulas | Aspose.Cells worksheet protection example
// Developer Intent: The developer needs to secure a spreadsheet so that users can modify only the cells that contain formulas while all other cells stay read‑only.
// Use Cases: Distribute a financial model where analysts can update calculation results but cannot alter input data or static text. | Create a reporting template that locks headings and constants, leaving only dynamic formula cells editable for end users. | Share a collaborative workbook where formula logic can be refined without risking accidental changes to raw values.
// AI Prompts: Generate C# code that protects an Aspose.Cells worksheet and unlocks only cells where IsFormula is true. | Explain how to iterate through a worksheet's used range, set Style.IsLocked based on Cell.IsFormula, and apply password protection. | Show an Aspose.Cells example that saves a workbook after protecting it so that only formula cells are editable.

using System;
using Aspose.Cells;

// Creates a workbook, unlocks cells that contain formulas, locks all other cells, applies worksheet protection with a password, and saves the file so only formula cells remain editable using Aspose.Cells for .NET.
class ProtectFormulaOnly
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
        cells["A3"].Formula = "=A1+A2";          // formula cell
        cells["B1"].PutValue("Sample Text");
        cells["B2"].Formula = "=NOW()";          // formula cell

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Iterate through all used cells
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];

                // Create a new style for the cell
                Style style = workbook.CreateStyle();

                if (cell.IsFormula)               // Allow editing of formula cells
                {
                    style.IsLocked = false;       // Unlock formula cells
                }
                else
                {
                    style.IsLocked = true;        // Keep other cells locked
                }

                // Apply the style to the cell
                cell.SetStyle(style);
            }
        }

        // Protect the worksheet (all protection types) with a password.
        // Only unlocked cells (the formula cells) can be edited.
        worksheet.Protect(ProtectionType.All, "myPassword", null);

        // Save the workbook
        workbook.Save("FormulaEditable.xlsx");
    }
}
