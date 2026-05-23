using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (first worksheet is at index 0)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a style that uses the Accent1 theme color as background
            Style accentStyle = workbook.CreateStyle();
            accentStyle.Pattern = BackgroundType.Solid;                     // Solid fill
            accentStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0); // No tint

            // Apply the style to a range of cells in the first worksheet
            // Example: cells A1 to D5
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    cell.PutValue($"R{row + 1}C{col + 1}");
                    cell.SetStyle(accentStyle);
                }
            }

            // Save the workbook
            workbook.Save("Accent1ThemeBackground.xlsx");
        }
    }
}