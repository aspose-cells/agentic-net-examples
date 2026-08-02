using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeSummary
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default theme will be used)
            Workbook workbook = new Workbook();

            // Add a new worksheet for the summary
            Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
            summarySheet.Name = "Theme Summary";

            // Write header titles
            summarySheet.Cells[0, 0].PutValue("Theme Color Name");
            summarySheet.Cells[0, 1].PutValue("Hexadecimal Value");

            // Iterate through all ThemeColorType values
            int rowIndex = 1; // start after header
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Retrieve the theme color from the workbook
                Color themeColor = workbook.GetThemeColor(type);

                // Convert the color to a hex string (RGB)
                string hex = $"#{themeColor.R:X2}{themeColor.G:X2}{themeColor.B:X2}";

                // Write the name and hex value to the sheet
                summarySheet.Cells[rowIndex, 0].PutValue(type.ToString());
                summarySheet.Cells[rowIndex, 1].PutValue(hex);

                rowIndex++;
            }

            // Save the workbook to a file
            workbook.Save("ThemeSummary.xlsx");
        }
    }
}