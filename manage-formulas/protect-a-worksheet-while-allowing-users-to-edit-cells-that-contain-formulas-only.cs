// Title: Protect an Excel Worksheet but Keep Formula Cells Editable with Aspose.Cells for .NET (C#)
// Description: Shows how to unlock only formula cells by setting Style.IsLocked = false, protect the entire worksheet with ProtectionType.All and a password, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells protect worksheet C# | unlock formula cells Aspose.Cells | Cell.IsFormula | Style.IsLocked | worksheet protection password | C# Excel lock unlock cells | Aspose.Cells example | editable formulas in protected sheet | Excel template formula editing
// Common Searches: Aspose.Cells protect sheet but allow formula editing | C# unlock cells with formulas before protecting worksheet | How to keep formula cells editable in a protected Excel file using Aspose.Cells | Set IsLocked false for formula cells Aspose.Cells .NET | Protect Excel worksheet with password using Aspose.Cells
// Developer Intent: Unlock only cells that contain formulas, then protect the whole worksheet with a password so users can edit those formulas while all other cells remain locked.
// Use Cases: Financial model where end‑users can modify calculated results but cannot change input constants. | Reporting dashboard that locks static data yet permits analysts to adjust formula outcomes. | Excel‑based data‑entry form that secures reference cells while allowing formula edits. | Template distribution that prevents changes to constants but keeps calculations editable.
// AI Prompts: Generate C# code using Aspose.Cells to iterate over used cells, set Style.IsLocked = false for cells where Cell.IsFormula is true, then protect the worksheet with ProtectionType.All and a password. | Explain step‑by‑step how Cell.IsFormula and Style.IsLocked work together to allow editing of formula cells on a protected sheet in Aspose.Cells. | Provide a complete Aspose.Cells .NET example that creates a workbook, adds constants and formulas, unlocks formula cells, protects the sheet, and saves the file.

using System;
using Aspose.Cells;

namespace WorksheetProtectionDemo
{
    // Shows how to unlock only formula cells by setting Style.IsLocked = false, protect the entire worksheet with ProtectionType.All and a password, and save the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: some constants and some formulas
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].Formula = "=A1+A2";          // formula cell
            cells["B1"].PutValue("Text");
            cells["B2"].Formula = "=LEN(B1)";        // formula cell
            cells["C1"].PutValue(5);
            cells["C2"].Formula = "=C1*2";           // formula cell

            // Unlock only the cells that contain formulas so they can be edited after protection
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.IsFormula)
                    {
                        // Get the current style, set IsLocked to false, and apply it back
                        Style style = cell.GetStyle();
                        style.IsLocked = false;   // unlocked cells can be edited on a protected sheet
                        cell.SetStyle(style);
                    }
                }
            }

            // Protect the worksheet (all protection types) with a password
            sheet.Protect(ProtectionType.All, "SecurePwd123", null);

            // Save the workbook
            workbook.Save("WorksheetProtected_FormulaEditable.xlsx");
        }
    }
}
