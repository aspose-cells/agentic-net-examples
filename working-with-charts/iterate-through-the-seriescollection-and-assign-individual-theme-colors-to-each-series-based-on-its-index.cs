// Title: C# – Assign Theme Colors to Each Chart Series Using Aspose.Cells SeriesCollection Loop
// Description: Creates a workbook, adds three data series, builds a column chart, defines an Accent1‑Accent6 palette, and iterates the SeriesCollection to apply a cyclic ThemeColor to each series border before saving the file.
// Keywords: Aspose.Cells C# chart series theme color | SeriesCollection loop Aspose.Cells | ThemeColorType Accent palette .NET | apply border color to chart series | column chart color customization | Aspose.Cells example USA | Aspose.Cells tutorial India | chart series styling C#
// Common Searches: How to set different theme colors for each series in an Aspose.Cells chart (C#) | C# loop through SeriesCollection to change chart series border color | Apply cyclic Accent colors to Aspose.Cells column chart series | Aspose.Cells chart series color palette example | Assign theme colors to chart series in .NET (US developers) | Aspose.Cells chart styling guide UK
// Developer Intent: Apply a distinct, cyclic theme color to every series in an Aspose.Cells chart via code.
// Use Cases: Clearly separate multiple data series in a column chart with corporate accent colors. | Automatically color an unknown number of series by reusing a predefined theme palette. | Maintain consistent branding across charts by using Aspose.Cells built‑in Accent1‑Accent6 colors.
// AI Prompts: Generate C# code that sets a ThemeColor fill format (instead of border) for each series using the same cyclic palette. | Provide a method to restore all series to the default Aspose.Cells theme after custom colors have been applied. | Explain how to map custom RGB values to chart series when the built‑in ThemeColorType palette is insufficient.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsSeriesThemeDemo
{
    // Creates a workbook, adds three data series, builds a column chart, defines an Accent1‑Accent6 palette, and iterates the SeriesCollection to apply a cyclic ThemeColor to each series border before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for multiple series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // Series 1
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Series 2
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Series 3
            sheet.Cells["D1"].PutValue("Series3");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series (B1:D4) and categories (A2:A4)
            chart.NSeries.Add("B1:D4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define a set of theme colors to apply (Accent1‑Accent6)
            ThemeColorType[] themePalette = new ThemeColorType[]
            {
                ThemeColorType.Accent1,
                ThemeColorType.Accent2,
                ThemeColorType.Accent3,
                ThemeColorType.Accent4,
                ThemeColorType.Accent5,
                ThemeColorType.Accent6
            };

            // Iterate through each series and assign a theme color based on its index
            SeriesCollection seriesColl = chart.NSeries;
            for (int i = 0; i < seriesColl.Count; i++)
            {
                // Choose a theme color cyclically from the palette
                ThemeColorType colorType = themePalette[i % themePalette.Length];

                // Apply the theme color to the series border (you could also set Area.FillFormat, etc.)
                seriesColl[i].Border.ThemeColor = new ThemeColor(colorType, 0.0);
                seriesColl[i].Border.Style = LineType.Solid;
                seriesColl[i].Border.Weight = WeightType.MediumLine;
            }

            // Save the workbook
            workbook.Save("SeriesCollection_ThemeColors_Demo.xlsx");
        }
    }
}
