using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Define the corporate green color
        Color corporateGreen = Color.FromArgb(0, 128, 0); // example corporate green

        // Replace the theme's Accent4 with the corporate green
        workbook.SetThemeColor(ThemeColorType.Accent4, corporateGreen);

        // Verify that chart series use the updated Accent4 theme color
        foreach (Worksheet ws in workbook.Worksheets)
        {
            foreach (Chart chart in ws.Charts)
            {
                if (chart.NSeries.Count == 0) continue;

                // Examine the first series (adjust as needed for multiple series)
                var series = chart.NSeries[0];
                var fillFormat = series.Area.FillFormat;

                // Ensure the fill is a solid fill to access CellsColor
                if (fillFormat.Type == FillType.Solid)
                {
                    CellsColor cellsColor = fillFormat.SolidFill.CellsColor;
                    ThemeColor themeColor = cellsColor.ThemeColor;

                    // Check if the series uses Accent4
                    if (themeColor != null && themeColor.ColorType == ThemeColorType.Accent4)
                    {
                        Console.WriteLine($"Chart '{chart.Name}' series correctly uses Accent4.");
                    }
                    else
                    {
                        Console.WriteLine($"Chart '{chart.Name}' series does NOT use Accent4.");
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}