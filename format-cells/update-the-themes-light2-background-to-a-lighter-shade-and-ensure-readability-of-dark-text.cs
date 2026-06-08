using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeUpdateExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Lighten the theme's Background2 (Light2) color
            // Using a light pastel color to make the background lighter
            workbook.SetThemeColor(ThemeColorType.Background2, Color.LightYellow);

            // Prepare a style that uses the updated Background2 theme color
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;
            // Apply the (now lighter) Background2 theme color
            style.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background2, 0.0);
            // Ensure dark text is readable: use Text1 theme color (typically dark) with no tint
            style.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0.0);
            style.Font.IsBold = true;
            style.Font.Size = 12;

            // Apply the style to a sample cell
            Cell cell = cells["A1"];
            cell.PutValue("Dark text on lighter Background2");
            cell.SetStyle(style);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("UpdatedThemeLight2.xlsx");
        }
    }
}