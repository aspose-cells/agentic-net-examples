using System;
using System.Drawing;
using Aspose.Cells;

class ApplyCellStyleDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Create a style using the workbook's factory method
        Style style = workbook.CreateStyle();

        // Set bold font
        style.Font.IsBold = true;

        // Set red font color
        style.Font.Color = Color.Red;

        // Apply thin black borders on all four sides
        style.SetBorder(BorderType.LeftBorder, CellBorderType.Thin, Color.Black);
        style.SetBorder(BorderType.RightBorder, CellBorderType.Thin, Color.Black);
        style.SetBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
        style.SetBorder(BorderType.BottomBorder, CellBorderType.Thin, Color.Black);

        // Get the target cell (B2) and put a sample value
        Cell cell = sheet.Cells[1, 1];
        cell.PutValue("Styled Text");

        // Apply the created style to the cell
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("StyledCell.xlsx", SaveFormat.Xlsx);
    }
}