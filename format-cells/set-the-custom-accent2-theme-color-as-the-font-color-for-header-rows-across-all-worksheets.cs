using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHeaderThemeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Define a custom color for the Accent2 theme (example: teal)
            Color customAccent2 = Color.FromArgb(0, 128, 128);

            // Apply the custom color to the Accent2 theme slot
            workbook.SetThemeColor(ThemeColorType.Accent2, customAccent2);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last column that contains data in the sheet
                int lastColumn = sheet.Cells.MaxDataColumn;

                // Apply the Accent2 theme color to every cell in the first row (header row)
                for (int col = 0; col <= lastColumn; col++)
                {
                    Cell headerCell = sheet.Cells[0, col];
                    Style style = headerCell.GetStyle();

                    // Set the font to use the Accent2 theme color (no tint)
                    style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0.0);

                    // Apply the modified style back to the cell
                    headerCell.SetStyle(style);
                }
            }

            // Save the workbook to a file
            workbook.Save("HeaderWithAccent2Theme.xlsx");
        }
    }
}