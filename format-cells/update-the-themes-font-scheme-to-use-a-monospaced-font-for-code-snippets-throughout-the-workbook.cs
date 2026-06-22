using System;
using Aspose.Cells;
using System.Drawing;

class UpdateThemeFontScheme
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Set the default style font to a monospaced font (e.g., Consolas)
        // This will affect cells that use the default style, effectively updating the theme's font scheme for code snippets
        Style defaultStyle = workbook.DefaultStyle;
        defaultStyle.Font.Name = "Consolas";          // Monospaced font
        defaultStyle.Font.Size = 11;                  // Typical size for code
        defaultStyle.Font.SchemeType = FontSchemeType.None; // Ensure explicit font name is used
        workbook.DefaultStyle = defaultStyle;

        // Optionally, apply the default style to all existing cells to ensure consistency
        Worksheet sheet = workbook.Worksheets[0];
        int maxRow = sheet.Cells.MaxDataRow;
        int maxCol = sheet.Cells.MaxDataColumn;
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = sheet.Cells[row, col];
                // Preserve existing value, apply default style
                Style cellStyle = cell.GetStyle();
                cellStyle.Font.Name = defaultStyle.Font.Name;
                cellStyle.Font.Size = defaultStyle.Font.Size;
                cellStyle.Font.SchemeType = defaultStyle.Font.SchemeType;
                cell.SetStyle(cellStyle);
            }
        }

        // Save the workbook (lifecycle rule: save)
        workbook.Save("UpdatedThemeFontScheme.xlsx");
    }
}