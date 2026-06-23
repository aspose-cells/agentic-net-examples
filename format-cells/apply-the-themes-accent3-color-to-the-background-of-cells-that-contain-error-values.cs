using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ApplyAccent3ToErrorCells
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add some formulas that will generate errors
        cells["A1"].Formula = "=1/0";                 // #DIV/0! error
        cells["A2"].Formula = "=UNKNOWNFUNC(1)";     // #NAME? error
        cells["A3"].PutValue(123);                   // Normal value

        // Calculate formulas to materialize errors (ignore exceptions)
        try { workbook.CalculateFormula(); } catch { }

        // Scan all used cells for error values
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int r = 0; r <= maxRow; r++)
        {
            for (int c = 0; c <= maxCol; c++)
            {
                Cell cell = cells[r, c];
                CellRichValue rich = cell.GetRichValue();

                // If the cell contains an error, apply Accent3 background theme color
                if (rich != null && rich.ErrorValue != default(ErrorCellValueType))
                {
                    Style style = workbook.CreateStyle();
                    style.Pattern = BackgroundType.Solid;
                    style.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);
                    cell.SetStyle(style);
                }
            }
        }

        // Save the workbook
        workbook.Save("ErrorCellsAccent3.xlsx");
    }
}