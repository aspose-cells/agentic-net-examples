using System;
using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a style and set the left border to red
        Style leftBorderStyle = workbook.CreateStyle();
        leftBorderStyle.Borders[BorderType.LeftBorder].Color = Color.Red;
        leftBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;

        // Create a style flag to apply only the left border
        StyleFlag flag = new StyleFlag();
        flag.LeftBorder = true;

        // Apply the style to column T (zero‑based index 19)
        cells.ApplyColumnStyle(19, leftBorderStyle, flag);

        // Save the workbook
        workbook.Save("ColumnT_LeftBorder.xlsx");
    }
}