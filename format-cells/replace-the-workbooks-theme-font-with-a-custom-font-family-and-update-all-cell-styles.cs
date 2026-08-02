using System;
using Aspose.Cells;
using System.Drawing;

class ReplaceThemeFont
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Define the custom font family to be used as the theme font
        string customFontFamily = "Arial";

        // Update the default style so that any new cells will inherit the custom font
        Style defaultStyle = workbook.DefaultStyle;
        defaultStyle.Font.Name = customFontFamily;
        // Apply the font as a major scheme font (you can also set Minor if needed)
        defaultStyle.Font.SchemeType = FontSchemeType.Major;
        workbook.DefaultStyle = defaultStyle;

        // Iterate through all worksheets and cells to replace existing font settings
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range to limit iteration (optional: you can use MaxDataRow/Column)
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    // Skip empty cells to avoid unnecessary style creation
                    if (cell == null || cell.Type == CellValueType.IsNull) continue;

                    // Get the current style, modify the font, and reapply
                    Style style = cell.GetStyle();
                    style.Font.Name = customFontFamily;
                    style.Font.SchemeType = FontSchemeType.Major;
                    cell.SetStyle(style);
                }
            }
        }

        // Save the workbook (lifecycle rule: save)
        workbook.Save("WorkbookWithCustomThemeFont.xlsx");
    }
}