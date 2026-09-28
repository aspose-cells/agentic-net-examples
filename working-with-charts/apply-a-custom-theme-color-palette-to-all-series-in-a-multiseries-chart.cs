// Title: How to apply a custom workbook theme and a monochromatic palette to every series in a multi‑series column chart using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, defines a 12‑color custom theme, assigns it to the workbook, adds a column chart with multiple series, and changes all series colors to the MonochromaticPalette1 palette. | Show the steps to use Aspose.Cells API to set a custom theme named "MyCustomTheme" and then call SeriesCollection.ChangeColors with ChartColorPaletteType.MonochromaticPalette1 for a chart containing several data series.
// Common Searches: aspnet apply custom theme to workbook and use it for chart series | aspnet cells change all series colors to monochromatic palette | c# set custom theme colors for Aspose.Cells chart | how to use SeriesCollection.ChangeColors in Aspose.Cells .NET | apply custom theme and monochrome palette to multi series column chart Aspose.Cells
// Tags: custom workbook theme Aspose.Cells .NET | monochromatic chart palette Aspose.Cells | SeriesCollection.ChangeColors usage | ChartColorPaletteType.MonochromaticPalette1 example | multi‑series column chart color customization

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomThemeDemo
{
    // Demonstrates creating a workbook, defining a 12‑color custom theme, applying it, adding a multi‑series column chart, and switching all series to a monochromatic palette with Aspose.Cells for .NET.
    public class ApplyCustomThemeToChartSeries
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a multi‑series chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // Series 1 data
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Series 2 data
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart that will contain both series
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series (both columns B and C)
            chart.NSeries.Add("B1:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ------------------------------------------------------------
            // 1. Define a custom theme (12 colors) and apply it to the workbook
            // ------------------------------------------------------------
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1
                Color.FromArgb(0, 0, 0),       // Text1
                Color.FromArgb(240, 240, 240), // Background2
                Color.FromArgb(80, 80, 80),    // Text2
                Color.FromArgb(0, 112, 192),   // Accent1
                Color.FromArgb(255, 192, 0),   // Accent2
                Color.FromArgb(112, 173, 71),  // Accent3
                Color.FromArgb(255, 0, 0),     // Accent4
                Color.FromArgb(0, 176, 80),    // Accent5
                Color.FromArgb(0, 176, 240),   // Accent6
                Color.FromArgb(0, 0, 255),     // Hyperlink
                Color.FromArgb(128, 0, 128)    // Followed Hyperlink
            };

            // Apply the custom theme to the workbook
            workbook.CustomTheme("MyCustomTheme", customColors);

            // ------------------------------------------------------------
            // 2. Apply a monochromatic palette to all series in the chart
            // ------------------------------------------------------------
            SeriesCollection seriesColl = chart.NSeries;
            seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette1);

            // ------------------------------------------------------------
            // Save the workbook to verify the result
            // ------------------------------------------------------------
            string outputPath = "CustomTheme_MultiSeriesChart.xlsx";
            workbook.Save(outputPath);
        }
    }
}
