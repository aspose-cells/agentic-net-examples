using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHeaderTheme
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Define the custom color you want for Accent2 (example: Orange)
            Color customAccent2 = Color.Orange;

            // Set the Accent2 theme color to the custom color
            workbook.SetThemeColor(ThemeColorType.Accent2, customAccent2);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last used column in the sheet
                int lastColumn = sheet.Cells.MaxDataColumn;

                // Apply the Accent2 theme color to each cell in the first row (header row)
                for (int col = 0; col <= lastColumn; col++)
                {
                    Cell headerCell = sheet.Cells[0, col]; // Row 0 is the header row

                    // Get the current style of the cell
                    Style style = headerCell.GetStyle();

                    // Set the font to use the Accent2 theme color (tint = 0 means no adjustment)
                    style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0);

                    // Apply the modified style back to the cell
                    headerCell.SetStyle(style);
                }
            }

            // Save the workbook to a file
            workbook.Save("HeaderWithAccent2Theme.xlsx");
        }
    }
}