// Title: C# – Assign Cyclic Theme Colors to Chart Series with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a column chart, and programmatically iterate the SeriesCollection to apply Accent theme colors to each series border. Colors are assigned cyclically, with a solid line style and medium weight, then the workbook is saved as an .xlsx file.
// Keywords: Aspose.Cells chart series theme color C# | Aspose.Cells set series border color | C# assign theme colors to chart series | Aspose.Cells SeriesCollection iteration | cycle chart series colors Aspose | ThemeColorType Accent1 Aspose.Cells | apply solid line style to chart series | Aspose.Cells column chart formatting | programmatic chart styling Aspose.Cells | C# Excel chart series color automation
// Common Searches: How to change chart series colors in Aspose.Cells C# | C# Aspose.Cells assign different theme colors to each series | Iterate SeriesCollection to set border color Aspose.Cells | Cyclic Accent theme colors for Excel chart series using Aspose | Set line style and weight for chart series programmatically Aspose.Cells
// Developer Intent: Programmatically give each chart series a unique theme color (cycling through a predefined list) and consistent border styling using Aspose.Cells for .NET.
// Use Cases: Generate multi‑series column charts where each series is visually distinct via Accent theme colors. | Apply uniform solid line style and medium weight to all series while varying colors automatically. | Create templates that support any number of series by reusing the six Accent colors cyclically. | Standardize chart appearance across reports generated in C# with Aspose.Cells.
// AI Prompts: Write C# code with Aspose.Cells that loops through all series in a chart and sets the fill color using ThemeColorType values, cycling through Accent1‑Accent6. | Show how to customize the border thickness and dash style for each series while still applying cyclic theme colors. | Create a function that accepts a custom list of ThemeColorType and applies them to chart series borders in Aspose.Cells. | Generate a complete example that reads series count from an existing chart and assigns theme colors without hard‑coding the list.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsSeriesThemeDemo
{
    // Demonstrates how to create a workbook, add a column chart, and programmatically iterate the SeriesCollection to apply Accent theme colors to each series border. Colors are assigned cyclically, with a solid line style and medium weight, then the workbook is saved as an .xlsx file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["D1"].PutValue("Series 3");

            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));
                sheet.Cells[$"B{i}"].PutValue(i * 10);
                sheet.Cells[$"C{i}"].PutValue(i * 12);
                sheet.Cells[$"D{i}"].PutValue(i * 14);
            }

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 15);
            Chart chart = sheet.Charts[chartIdx];

            // Set chart data range (including all series)
            chart.NSeries.Add("B2:D6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Define a list of theme colors to assign (using Accent theme colors)
            List<ThemeColorType> themeColors = new List<ThemeColorType>
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6
            };

            // Iterate through each series in the collection and assign a theme color based on its index
            SeriesCollection seriesColl = chart.NSeries;
            for (int i = 0; i < seriesColl.Count; i++)
            {
                // Choose a theme color cyclically
                ThemeColorType colorType = themeColors[i % themeColors.Count];

                // Apply the theme color to the series border (you can also set Fill, etc.)
                seriesColl[i].Border.ThemeColor = new ThemeColor(colorType, 0.0);
                seriesColl[i].Border.Style = LineType.Solid;
                seriesColl[i].Border.Weight = WeightType.MediumLine;
            }

            // Save the workbook
            workbook.Save("SeriesCollection_ThemeColors.xlsx");
        }
    }
}
