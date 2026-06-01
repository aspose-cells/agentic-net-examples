using System;
using Aspose.Cells;

namespace AsposeCellsLockFormulas
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data: input cells
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].PutValue(5);
            cells["B2"].PutValue(15);

            // Sample formulas that should be locked
            cells["C1"].Formula = "=A1+B1";
            cells["C2"].Formula = "=A2+B2";

            // Calculate formulas so that results are available
            workbook.CalculateFormula();

            // Iterate through all used cells and set lock status
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    Style style = cell.GetStyle();

                    // Lock if the cell contains a formula, otherwise unlock
                    style.IsLocked = !string.IsNullOrEmpty(cell.Formula);
                    cell.SetStyle(style);
                }
            }

            // Protect the worksheet (all protection types) with a password
            worksheet.Protect(ProtectionType.All, "securePassword", null);

            // Save the workbook
            workbook.Save("LockedFormulas.xlsx");
        }
    }
}