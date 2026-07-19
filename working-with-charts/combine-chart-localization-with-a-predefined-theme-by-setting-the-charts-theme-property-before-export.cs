// Title: C# – Apply a Predefined Theme and Localization to an Aspose.Cells Chart and Export to PDF
// Description: This example shows how to create a workbook, define a custom theme with specific colors, assign the theme to the workbook, configure ChartGlobalizationSettings for localized titles and axis labels, set the chart's Theme property, add a column chart, apply theme colors to the series, and export the chart to a PDF while also saving the workbook.
// Keywords: Aspose.Cells C# chart theme | custom theme Aspose.Cells | chart localization Aspose.Cells | ChartGlobalizationSettings .NET | export chart to PDF Aspose.Cells | Workbook.CustomTheme example | theme colors for chart series | multilingual chart titles | apply predefined theme to chart | Aspose.Cells chart PDF export
// Common Searches: How to set a custom theme for a chart in Aspose.Cells C# | Aspose.Cells chart localization with globalization settings | Export a themed chart to PDF using Aspose.Cells | Apply predefined theme colors to Aspose.Cells chart series | Combine Workbook.CustomTheme and ChartGlobalizationSettings
// Developer Intent: Create a localized chart that uses a predefined theme and export it as a PDF.
// Use Cases: Generate multilingual financial reports where charts follow corporate branding. | Produce PDF dashboards that maintain a consistent color palette across all visualizations. | Reuse a custom theme in multiple workbooks while automatically translating chart labels.
// AI Prompts: Show me C# code to define a custom theme, assign it to a workbook, set the chart's Theme property, and export the chart to PDF with localized titles. | Provide an example of using ChartGlobalizationSettings in Aspose.Cells to translate chart titles and axis labels, then save the chart as a PDF. | Explain how Workbook.CustomTheme and Chart.Theme work together to create a themed, localized chart in Aspose.Cells .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartThemeLocalization
{
    // Custom globalization settings for chart localization
    // This example shows how to create a workbook, define a custom theme with specific colors, assign the theme to the workbook, configure ChartGlobalizationSettings for localized titles and axis labels, set the chart's Theme property, add a column chart, apply theme colors to the series, and export the chart to a PDF while also saving the workbook.
    public class MyChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetChartTitleName() => "Localized Chart Title";
        public override string GetAxisTitleName() => "Localized Axis Title";
        public override string GetSeriesName() => "Localized Series";
        // Other overrides can be added as needed
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Apply a predefined custom theme ----------
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
            workbook.CustomTheme("MyPredefinedTheme", customColors);

            // ---------- Set chart localization (globalization) ----------
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new MyChartGlobalizationSettings()
            };

            // ---------- Populate sample data ----------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // ---------- Add a chart ----------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Use the theme colors for the series (e.g., Accent1)
            chart.NSeries[0].Border.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.6);
            chart.NSeries[0].Border.Style = LineType.Solid;
            chart.NSeries[0].Border.Weight = WeightType.MediumLine;

            // Set a localized title (will be overridden by globalization if needed)
            chart.Title.Text = "Demo Chart";

            // ---------- Export the chart to PDF (chart respects the theme) ----------
            chart.ToPdf("LocalizedChartWithTheme.pdf");

            // Save the workbook for reference
            workbook.Save("LocalizedChartWithTheme.xlsx");
        }
    }
}
