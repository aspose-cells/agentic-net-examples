// Title: Aspose.Cells C# – Override ChartGlobalizationSettings to Localize Chart Legends to Arabic
// Description: Demonstrates how to create a custom ArabicChartGlobalizationSettings class that inherits from ChartGlobalizationSettings and overrides legend, series, title, axis and unit methods to return Arabic strings. The custom settings are assigned to Workbook.Settings.GlobalizationSettings, so every chart in the workbook automatically displays Arabic labels such as "زيادة", "نقصان", and "المجموع".
// Keywords: Aspose.Cells | C# | ChartGlobalizationSettings | Arabic chart localization | Excel legend translation | custom chart globalization | Arabic Excel charts | DisplayUnitType Arabic | Middle East localization | UAE Excel reporting
// Common Searches: How to localize Aspose.Cells chart legends to Arabic | Override ChartGlobalizationSettings for Arabic labels in .NET | Aspose.Cells Arabic chart title and axis names | Set global Arabic chart settings in a workbook | C# example for Arabic Excel chart legends
// Developer Intent: Provide Arabic translations for all chart legend entries and related text by customizing ChartGlobalizationSettings.
// Use Cases: Generate Excel reports for Arabic‑speaking audiences with automatically translated chart legends, titles, and axis labels. | Apply a single globalization setting to ensure every chart in a workbook uses Arabic terminology without per‑chart code. | Create localized financial dashboards where terms like "Increase", "Decrease", and "Total" appear in Arabic.
// AI Prompts: Write C# code that overrides all ChartGlobalizationSettings legend methods to return Arabic strings and applies the settings to a workbook. | Extend ArabicChartGlobalizationSettings to include Arabic names for DisplayUnitType values such as Millions and Billions. | Show how to use ArabicChartGlobalizationSettings with a pie chart and verify Arabic legend entries.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsArabicChartDemo
{
    // Custom globalization settings that provide Arabic translations for legend entries
    // Demonstrates how to create a custom ArabicChartGlobalizationSettings class that inherits from ChartGlobalizationSettings and overrides legend, series, title, axis and unit methods to return Arabic strings. The custom settings are assigned to Workbook.Settings.GlobalizationSettings, so every chart in the workbook automatically displays Arabic labels such as "زيادة", "نقصان", and "المجموع".
    public class ArabicChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // Arabic text for the "Increase" legend entry
        public override string GetLegendIncreaseName()
        {
            return "زيادة";
        }

        // Arabic text for the "Decrease" legend entry
        public override string GetLegendDecreaseName()
        {
            return "نقصان";
        }

        // Arabic text for the "Total" legend entry
        public override string GetLegendTotalName()
        {
            return "المجموع";
        }

        // Optionally override other methods to provide Arabic equivalents
        public override string GetSeriesName()
        {
            return "سلسلة";
        }

        public override string GetChartTitleName()
        {
            return "عنوان المخطط";
        }

        public override string GetOtherName()
        {
            return "أخرى";
        }

        public override string GetAxisTitleName()
        {
            return "عنوان المحور";
        }

        public override string GetAxisUnitName(DisplayUnitType type)
        {
            // Example: Arabic unit names for common display units
            switch (type)
            {
                case DisplayUnitType.Hundreds:
                    return "مئات";
                case DisplayUnitType.Thousands:
                    return "آلاف";
                case DisplayUnitType.TenThousands:
                    return "عشرات الآلاف";
                default:
                    return base.GetAxisUnitName(type);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("الفئة");
            sheet.Cells["A2"].PutValue("الفئة 1");
            sheet.Cells["A3"].PutValue("الفئة 2");
            sheet.Cells["A4"].PutValue("الفئة 3");

            sheet.Cells["B1"].PutValue("القيمة");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "مخطط تجريبي";

            // Apply the custom Arabic globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ArabicChartGlobalizationSettings()
            };

            // Save the workbook
            workbook.Save("ArabicChartDemo.xlsx");
        }
    }
}
