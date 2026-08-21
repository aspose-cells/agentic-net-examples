// Title: Apply a Custom Theme and Localized Text to an Aspose.Cells Chart (C#)
// Description: This example shows how to create a workbook, define a 12‑color custom theme with Workbook.CustomTheme, configure chart globalization via a ChartGlobalizationSettings subclass, add sample data, insert a column chart, assign the data range, set the chart's Theme property, and save the file so the chart displays the corporate colors and localized titles/axis labels.
// Keywords: Aspose.Cells | C# | custom chart theme | chart localization | ChartGlobalizationSettings | Workbook.CustomTheme | chart Theme property | multilingual Excel chart | branding colors Excel | Aspose.Cells example
// Common Searches: Aspose.Cells set chart theme C# | How to localize chart titles with Aspose.Cells | Custom Excel theme for charts using Aspose | Apply corporate colors to Aspose.Cells chart | ChartGlobalizationSettings sample code | Export chart with localized strings Aspose.Cells
// Developer Intent: Create an Excel workbook where a chart uses a predefined custom color theme and localized titles/axis labels, then export the workbook.
// Use Cases: Produce sales dashboards that follow company branding while showing market‑specific language. | Generate multilingual financial reports where chart labels are automatically translated but retain a consistent visual style. | Automate recurring Excel exports with charts that always apply the same custom theme and localized text for global audiences. | Build Excel templates for regional teams that need both corporate colors and localized chart terminology.
// AI Prompts: Write C# code with Aspose.Cells to define a 12‑color custom theme, apply it to a workbook, set ChartGlobalizationSettings for localized titles, and save the chart. | Show how to subclass ChartGlobalizationSettings to provide custom chart titles and axis names, then combine it with a workbook theme for a column chart. | Explain step‑by‑step how to assign a custom theme to a chart via the Theme property, add localization, and export the workbook so the chart renders with the correct colors and text.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartThemeLocalization
{
    // Custom globalization settings for chart localization
    // This example shows how to create a workbook, define a 12‑color custom theme with Workbook.CustomTheme, configure chart globalization via a ChartGlobalizationSettings subclass, add sample data, insert a column chart, assign the data range, set the chart's Theme property, and save the file so the chart displays the corporate colors and localized titles/axis labels.
    public class MyChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetChartTitleName() => "Localized Chart Title";
        public override string GetAxisTitleName() => "Localized Axis Title";
        public override string GetSeriesName() => "Localized Series";
        // Other overrides can be added as needed
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Define a custom theme (12 colors) ----------
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

            // Apply the custom theme to the workbook (theme rule)
            workbook.CustomTheme("MyCustomTheme", customColors);

            // ---------- Apply chart localization ----------
            var chartSettings = new MyChartGlobalizationSettings();
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = chartSettings
            };

            // ---------- Populate sample data ----------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // ---------- Add a chart ----------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Apply a built‑in style (optional)
            chart.Style = 2; // Built‑in style index

            // The chart will now use the custom theme colors and the localized strings
            // Save the workbook (save rule)
            workbook.Save("ChartWithThemeAndLocalization.xlsx");
        }
    }
}
