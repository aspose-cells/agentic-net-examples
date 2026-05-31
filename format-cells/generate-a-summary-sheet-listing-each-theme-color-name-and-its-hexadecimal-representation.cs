using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeSummary
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet for the summary
            Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
            summarySheet.Name = "Theme Summary";

            // Write header titles
            summarySheet.Cells["A1"].PutValue("Theme Color Name");
            summarySheet.Cells["B1"].PutValue("Hexadecimal Value");

            // Get all ThemeColorType values
            ThemeColorType[] themeTypes = (ThemeColorType[])Enum.GetValues(typeof(ThemeColorType));

            // Start writing data from the second row
            int rowIndex = 1; // zero‑based index, row 2 in Excel

            foreach (ThemeColorType type in themeTypes)
            {
                // Skip the StyleColor entry as it is not a real theme color
                if (type == ThemeColorType.StyleColor) continue;

                // Retrieve the theme color from the workbook
                Color color = workbook.GetThemeColor(type);

                // Convert the color to a hex string (e.g., #RRGGBB)
                string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";

                // Write the theme name and its hex value to the sheet
                summarySheet.Cells[rowIndex, 0].PutValue(type.ToString()); // Column A
                summarySheet.Cells[rowIndex, 1].PutValue(hex);            // Column B

                rowIndex++;
            }

            // Auto‑fit columns for better readability
            summarySheet.AutoFitColumns();

            // Save the workbook
            workbook.Save("ThemeSummary.xlsx");
        }
    }
}