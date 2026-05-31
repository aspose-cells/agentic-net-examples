using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomThemeDemo
{
    public class ApplyCustomThemeToChartSeries
    {
        public static void Main()
        {
            try
            {
                Run();
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

            // Add a column chart that will contain the two series
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series (both columns B and C)
            chart.NSeries.Add("B1:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // -----------------------------------------------------------------
            // Define a custom theme (12 colors) that will be used by the chart
            // -----------------------------------------------------------------
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

            // ---------------------------------------------------------------
            // Apply a monochromatic palette to all series in the chart.
            // This uses the SeriesCollection.ChangeColors method.
            // ---------------------------------------------------------------
            SeriesCollection seriesColl = chart.NSeries;
            seriesColl.ChangeColors(ChartColorPaletteType.MonochromaticPalette1);

            // Save the workbook
            string outputPath = "CustomTheme_MultiSeriesChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}