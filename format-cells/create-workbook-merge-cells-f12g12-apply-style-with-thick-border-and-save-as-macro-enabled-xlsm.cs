using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Merge cells F12:G12 (zero‑based indices: row 11, column 5, 1 row, 2 columns)
        cells.Merge(11, 5, 1, 2);

        // Create a style with a thick black border on all sides
        Style thickBorderStyle = workbook.CreateStyle();
        thickBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
        thickBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
        thickBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
        thickBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
        thickBorderStyle.Borders[BorderType.TopBorder].Color = Color.Black;
        thickBorderStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
        thickBorderStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
        thickBorderStyle.Borders[BorderType.RightBorder].Color = Color.Black;

        // Apply the style to the merged cell (reference the upper‑left cell)
        cells[11, 5].SetStyle(thickBorderStyle);

        // Enable macros in the workbook (required for XLSM)
        workbook.Settings.EnableMacros = true;

        // Save the workbook as a macro‑enabled XLSM file
        workbook.Save("MergedCellWithBorder.xlsm", SaveFormat.Xlsm);
    }
}