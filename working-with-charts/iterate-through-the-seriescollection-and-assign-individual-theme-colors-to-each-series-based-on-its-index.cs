// Title: Aspose.Cells .NET – Assign Cyclic Theme Colors to Chart Series
// Description: Learn how to iterate a chart's SeriesCollection in Aspose.Cells for .NET and apply a distinct ThemeColor (Accent1‑Accent6) to each series. The example shows creating a workbook, adding a column chart, defining a reusable color list, and setting border style, weight, and fill in a cyclic manner before saving the file.
// Keywords: Aspose.Cells chart series theme color | C# assign accent colors to chart series | cycle series colors Aspose.Cells | SeriesCollection ThemeColor .NET | chart series border style Aspose
// Common Searches: set individual theme colors for chart series Aspose.Cells | apply accent colors to multiple series .NET | change chart series border color with ThemeColor | Aspose.Cells cyclic color assignment for series | customize chart series appearance Aspose
// Developer Intent: Apply a unique (or cyclic) theme color to every series in a chart generated with Aspose.Cells for .NET.
// Use Cases: Visually differentiate each data series in a column or line chart using corporate accent colors. | Map business‑specific ThemeColorType values to series for brand‑consistent reporting. | Handle charts with more series than available accents by looping through a predefined color palette.
// AI Prompts: Generate C# code that loops through a chart's SeriesCollection in Aspose.Cells and assigns ThemeColorType.Accent1‑Accent6 to each series border, cycling when necessary. | Show how to set both border and fill ThemeColor for chart series using a custom list of ThemeColorType values in Aspose.Cells .NET. | Explain how to customize line style, weight, and transparency when applying theme colors to chart series with Aspose.Cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsSeriesThemeDemo
{
    // Learn how to iterate a chart's SeriesCollection in Aspose.Cells for .NET and apply a distinct ThemeColor (Accent1‑Accent6) to each series. The example shows creating a workbook, adding a column chart, defining a reusable color list, and setting border style, weight, and fill in a cyclic manner before saving the file.
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
                sheet.Cells[$"A{i}"].PutValue($"Cat {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
                sheet.Cells[$"C{i}"].PutValue(i * 12);
                sheet.Cells[$"D{i}"].PutValue(i * 14);
            }

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 15);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series (B1:D6) and categories (A2:A6)
            chart.NSeries.Add("B1:D6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Define a list of theme colors to assign (Accent1‑Accent6)
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
                // Choose a theme color cyclically if there are more series than accents
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
