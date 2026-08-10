// Title: Assign ChartChineseSettings to GlobalizationSettings for Chinese Chart Localization in Aspose.Cells .NET
// Description: Demonstrates how to create a custom ChartChineseSettings class that inherits from ChartGlobalizationSettings, assign it to the workbook's GlobalizationSettings.ChartSettings, and export the workbook so the chart displays Chinese axis units and title.
// Keywords: Aspose.Cells chart Chinese localization | ChartGlobalizationSettings custom class | ChartChineseSettings example | set chart globalization settings .NET | override axis unit names Aspose.Cells | Chinese chart title Aspose.Cells | globalizationsettings chart aspnet
// Common Searches: Aspose.Cells custom chart globalization | How to apply Chinese settings to a chart in Aspose.Cells | Assign ChartGlobalizationSettings before saving workbook | Override chart axis labels in Chinese with Aspose.Cells | Set chart title localization Aspose.Cells .NET
// Developer Intent: Apply a custom Chinese globalization configuration to a chart by assigning a ChartGlobalizationSettings‑derived object to the workbook's GlobalizationSettings before exporting.
// Use Cases: Display axis units (hundreds, thousands, ten‑thousands) in Chinese characters. | Show a localized Chinese chart title without manual text replacement. | Ensure consistent Chinese formatting for multiple charts via workbook‑level settings.
// AI Prompts: Generate C# code that creates a ChartChineseSettings class, assigns it to Workbook.Settings.GlobalizationSettings.ChartSettings, and saves the file using Aspose.Cells. | Show how to override legend and series names in a custom ChartGlobalizationSettings class for Chinese localization. | Explain how to apply different ChartGlobalizationSettings objects to separate charts within the same workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to create a custom ChartChineseSettings class that inherits from ChartGlobalizationSettings, assign it to the workbook's GlobalizationSettings.ChartSettings, and export the workbook so the chart displays Chinese axis units and title.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate some sample data for the chart
        ws.Cells["A1"].PutValue("类别");
        ws.Cells["A2"].PutValue("第一组");
        ws.Cells["A3"].PutValue("第二组");
        ws.Cells["B1"].PutValue("数值");
        ws.Cells["B2"].PutValue(120);
        ws.Cells["B3"].PutValue(250);

        // Add a column chart to the worksheet
        int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = ws.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);          // Values
        chart.NSeries.CategoryData = "A2:A3";      // Categories

        // Apply custom Chinese globalization settings to the chart via workbook settings
        wb.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = new ChartChineseSettings()
        };

        // Export the workbook (the chart will use the Chinese settings)
        wb.Save("ChartChineseSettingsDemo.xlsx");
    }

    // Custom globalization settings for charts (Chinese locale)
    public class ChartChineseSettings : ChartGlobalizationSettings
    {
        // Override axis unit names (e.g., hundreds, thousands, ten‑thousands)
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            switch (type)
            {
                case DisplayUnitType.Hundreds:
                    return "百";
                case DisplayUnitType.Thousands:
                    return "千";
                case DisplayUnitType.TenThousands:
                    return "万";
                default:
                    return base.GetAxisUnitName(type);
            }
        }

        // Override the default chart title text
        public override string GetChartTitleName()
        {
            return "图表标题";
        }

        // Additional overrides can be added here as needed (e.g., legend labels, series names, etc.)
    }
}
