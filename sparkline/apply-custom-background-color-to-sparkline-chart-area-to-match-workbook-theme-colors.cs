using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineBackgroundDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the sparkline (A1:A4)
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["A2"].PutValue(3);
                sheet.Cells["A3"].PutValue(8);
                sheet.Cells["A4"].PutValue(2);

                // Define the cell where the sparkline will be placed (E1)
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4,
                    EndColumn = 4
                };

                // Add a sparkline group of type Line using the data range A1:A4
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A4", false, location);
                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // (Optional) Add additional sparklines if needed:
                // group.Sparklines.Add("A1:A4", 0, 4);

                // Retrieve a theme accent color (Accent1) from the workbook's theme
                Color themeAccentColor = workbook.GetThemeColor(ThemeColorType.Accent1);

                // Apply the theme color as the background of the sparkline cell (E1)
                Style style = sheet.Cells[0, 4].GetStyle();
                style.ForegroundColor = themeAccentColor;   // Set fill color
                style.Pattern = BackgroundType.Solid;       // Ensure solid fill
                sheet.Cells[0, 4].SetStyle(style);

                // Save the workbook
                workbook.Save("SparklineBackgroundThemeColor.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}