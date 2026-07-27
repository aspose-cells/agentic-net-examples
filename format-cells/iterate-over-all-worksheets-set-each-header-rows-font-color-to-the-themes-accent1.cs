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

            // Example: add some data and a header row to each worksheet for demonstration
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Populate header row (row 0) with sample column names
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[0, col].PutValue($"Header{col + 1}");
                }

                // Populate some data rows
                for (int row = 1; row <= 10; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
                    }
                }
            }

            // Retrieve the theme color for Accent1 from the workbook's theme
            Color accent1Color = workbook.GetThemeColor(ThemeColorType.Accent1);

            // Create a style that uses the Accent1 theme color for the font
            Style headerStyle = workbook.CreateStyle();
            // Option 1: set the actual color
            headerStyle.Font.Color = accent1Color;
            // Option 2 (alternative): use the theme color directly
            // headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);

            // Iterate over all worksheets and apply the style to the first (header) row
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the number of used columns in the sheet
                int maxColumn = sheet.Cells.MaxDataColumn;

                // Apply the style to each cell in the header row
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell headerCell = sheet.Cells[0, col];
                    headerCell.SetStyle(headerStyle);
                }
            }

            // Save the workbook
            workbook.Save("HeaderWithAccent1Theme.xlsx");
        }
    }
}