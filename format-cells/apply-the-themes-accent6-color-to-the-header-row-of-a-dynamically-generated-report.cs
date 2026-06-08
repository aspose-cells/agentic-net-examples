using System;
using System.Drawing;
using Aspose.Cells;

namespace ReportWithAccent6Header
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data for the report (5 columns, 10 rows)
            int rows = 10;
            int cols = 5;

            // Populate header row
            for (int c = 0; c < cols; c++)
            {
                cells[0, c].PutValue($"Header {c + 1}");
            }

            // Populate data rows
            for (int r = 1; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    cells[r, c].PutValue($"R{r}C{c + 1}");
                }
            }

            // Create a style that uses the theme's Accent6 color as background
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Pattern = BackgroundType.Solid;
            // Use Accent6 with no tint (0) – this will pick the theme's Accent6 color
            headerStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent6, 0);
            // Optional: make the font bold and white for contrast
            headerStyle.Font.IsBold = true;
            headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0); // usually white/black based on theme

            // Apply the style to the entire header row
            for (int c = 0; c < cols; c++)
            {
                cells[0, c].SetStyle(headerStyle);
            }

            // Save the workbook
            workbook.Save("ReportWithAccent6Header.xlsx");
        }
    }
}