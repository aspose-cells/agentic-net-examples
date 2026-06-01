using System;
using Aspose.Cells;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Create a style and set the left border to red and thin
        Style style = workbook.CreateStyle();
        style.Borders[BorderType.LeftBorder].Color = Color.Red;
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;

        // Create a style flag to apply only the left border
        StyleFlag flag = new StyleFlag();
        flag.LeftBorder = true;

        // Apply the style to column T (zero‑based index 19)
        cells.ApplyColumnStyle(19, style, flag);

        // Save the workbook
        workbook.Save("ColumnT_LeftBorder.xlsx");
    }
}