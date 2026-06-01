using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsThemeAndLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Apply a custom theme to the workbook
            // -------------------------------------------------
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1
                Color.FromArgb(0, 0, 0),       // Text1
                Color.FromArgb(240, 240, 240), // Background2
                Color.FromArgb(80, 80, 80),    // Text2
                Color.FromArgb(0, 112, 192),   // Accent1
                Color.FromArgb(255, 192, 0),   // Accent2
                Color.FromArgb(112, 48, 160),  // Accent3
                Color.FromArgb(0, 176, 80),    // Accent4
                Color.FromArgb(255, 0, 0),     // Accent5
                Color.FromArgb(0, 176, 240),   // Accent6
                Color.FromArgb(0, 0, 255),     // Hyperlink
                Color.FromArgb(128, 0, 128)    // Followed Hyperlink
            };
            workbook.CustomTheme("MyCustomTheme", customColors);

            // -------------------------------------------------
            // 2. Configure chart globalization (localization) settings
            // -------------------------------------------------
            var chartSettings = new SettableChartGlobalizationSettings();
            chartSettings.SetChartTitleName("売上高");                     // Chart title
            chartSettings.SetAxisTitleName("軸タイトル");                // Axis title
            chartSettings.SetSeriesName("シリーズ");                     // Series name
            chartSettings.SetLegendIncreaseName("増加");                // Legend increase
            chartSettings.SetLegendDecreaseName("減少");                // Legend decrease
            chartSettings.SetLegendTotalName("合計");                   // Legend total
            chartSettings.SetOtherName("その他");                       // "Other" label
            chartSettings.SetAxisUnitName(DisplayUnitType.Thousands, "千"); // Axis unit for thousands

            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = chartSettings
            };

            // -------------------------------------------------
            // 3. Populate sample data for the chart
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["B4"].PutValue(1800);

            // -------------------------------------------------
            // 4. Create a chart and apply the data
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories
            chart.Title.Text = "Sales Chart";

            // -------------------------------------------------
            // 5. Export the workbook (chart inherits the theme and localization)
            // -------------------------------------------------
            workbook.Save("ChartWithThemeAndLocalization.pdf");
        }
    }
}