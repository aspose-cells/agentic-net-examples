// Title: C# – Custom ChartChineseSettings subclass for Chinese localization of Aspose.Cells charts
// Description: Demonstrates how to create a ChartChineseSettings class that inherits from ChartGlobalizationSettings, overrides methods to supply Chinese strings for axis titles, display units, chart title, legend items, "Other" category and series name, and applies the settings to a workbook containing a column chart.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | Chinese localization | .NET | C# chart example | custom chart globalization | axis title translation | display unit Chinese | chart legend Chinese | globalization settings demo | GitHub Aspose.Cells example
// Common Searches: Aspose.Cells chart Chinese labels | How to localize chart axis titles in Chinese using Aspose.Cells | Custom ChartGlobalizationSettings C# example | Set Chinese display unit names in Aspose.Cells chart | ChartChineseSettings tutorial
// Developer Intent: Create a subclass of ChartGlobalizationSettings that returns Chinese text for all chart UI elements and apply it to a workbook.
// Use Cases: Generate Excel reports with fully Chinese chart captions for Chinese‑speaking audiences. | Show custom unit symbols such as “千” for thousands on chart axes. | Provide Chinese legends, titles, and "Other" category labels in pie or column charts.
// AI Prompts: Write a ChartGlobalizationSettings subclass that returns Japanese translations for chart elements. | Show how to toggle between English and Chinese chart globalization settings at runtime in Aspose.Cells. | Explain how to customize chart legend text using a custom globalization class in Aspose.Cells .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Custom globalization settings for charts providing Chinese translations
    // Demonstrates how to create a ChartChineseSettings class that inherits from ChartGlobalizationSettings, overrides methods to supply Chinese strings for axis titles, display units, chart title, legend items, "Other" category and series name, and applies the settings to a workbook containing a column chart.
    public class ChartChineseSettings : ChartGlobalizationSettings
    {
        // Axis title (e.g., X Axis, Y Axis)
        public override string GetAxisTitleName()
        {
            return "轴标题";
        }

        // Unit names for display units
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
                case DisplayUnitType.Millions:
                    return "百万";
                case DisplayUnitType.Billions:
                    return "十亿";
                default:
                    return base.GetAxisUnitName(type);
            }
        }

        // Chart title
        public override string GetChartTitleName()
        {
            return "图表标题";
        }

        // Legend labels
        public override string GetLegendDecreaseName()
        {
            return "递减";
        }

        public override string GetLegendIncreaseName()
        {
            return "递增";
        }

        public override string GetLegendTotalName()
        {
            return "总计";
        }

        // Name for "Other" category in pie charts etc.
        public override string GetOtherName()
        {
            return "其他";
        }

        // Series name
        public override string GetSeriesName()
        {
            return "系列";
        }
    }

    public class GlobalizationSettingsDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("类别");
            sheet.Cells["A2"].PutValue("第一季度");
            sheet.Cells["A3"].PutValue("第二季度");
            sheet.Cells["A4"].PutValue("第三季度");
            sheet.Cells["B1"].PutValue("数值");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(2500);
            sheet.Cells["B4"].PutValue(1800);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "销售额";

            // Set display unit to thousands to see custom unit name
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
            chart.ValueAxis.IsDisplayUnitLabelShown = true;

            // Apply the custom Chinese chart globalization settings
            wb.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChartChineseSettings()
            };

            // Output some localized strings to console for verification
            Console.WriteLine("Chart Title Name: " + ((ChartGlobalizationSettings)wb.Settings.GlobalizationSettings.ChartSettings).GetChartTitleName());
            Console.WriteLine("Axis Unit Name (Thousands): " + ((ChartGlobalizationSettings)wb.Settings.GlobalizationSettings.ChartSettings).GetAxisUnitName(DisplayUnitType.Thousands));
            Console.WriteLine("Legend Total Name: " + ((ChartGlobalizationSettings)wb.Settings.GlobalizationSettings.ChartSettings).GetLegendTotalName());

            // Save the workbook (the chart will reflect Chinese labels where applicable)
            wb.Save("ChartChineseSettingsDemo.xlsx");
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            GlobalizationSettingsDemo.Run();
        }
    }
}
