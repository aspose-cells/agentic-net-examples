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

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the last column that contains data in the first row (header row)
                int lastColumn = sheet.Cells.MaxDataColumn;

                // Loop through each cell in the header row (row index 0)
                for (int col = 0; col <= lastColumn; col++)
                {
                    // Get the cell at row 0, current column
                    Cell headerCell = sheet.Cells[0, col];

                    // Retrieve the current style of the cell
                    Style style = headerCell.GetStyle();

                    // Set the font to use the workbook's theme Accent1 color
                    style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);

                    // Create a StyleFlag to indicate that only the font color should be applied
                    StyleFlag flag = new StyleFlag();
                    flag.FontColor = true;

                    // Apply the style with the flag to the header cell
                    headerCell.SetStyle(style, flag);
                }
            }

            // Save the modified workbook
            workbook.Save("HeaderThemeAccent1.xlsx");
        }
    }
}