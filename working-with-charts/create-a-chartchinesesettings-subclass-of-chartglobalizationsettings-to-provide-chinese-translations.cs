// Title: C# – Localize Aspose.Cells Chart Labels to Chinese with a ChartChineseSettings Subclass
// Description: Demonstrates how to subclass ChartGlobalizationSettings, override methods to return Chinese strings for axis titles, unit names, legend and series labels, apply the settings to a workbook, and see the chart labels switch from English to Chinese before saving the file.
// Keywords: Aspose.Cells | C# | ChartGlobalizationSettings | Chinese localization | Excel chart translation | DisplayUnitType | globalization settings | chart axis label | ChartChineseSettings | Aspose.Cells .NET example
// Common Searches: Aspose.Cells chart Chinese translation example | ChartGlobalizationSettings subclass for Chinese language | C# change Excel chart axis unit label to Chinese | How to apply custom chart globalization settings in Aspose.Cells | Localize Aspose.Cells chart text for Chinese reports
// Developer Intent: Create a custom ChartGlobalizationSettings subclass that supplies Chinese text for all chart UI elements and attach it to a workbook’s globalization settings.
// Use Cases: Provide Chinese axis titles, unit names, chart title, legend labels, and series names for localized Excel dashboards. | Automatically display Chinese unit labels (百, 千, 万, 百万, 十亿) when the chart’s DisplayUnitType changes. | Reuse the ChartChineseSettings class across multiple workbooks to ensure consistent Chinese chart terminology. | Generate Excel files with fully Chinese‑localized charts for business intelligence or reporting in China.
// AI Prompts: Write C# code that defines a ChartChineseSettings class inheriting from ChartGlobalizationSettings and overrides all relevant methods to return Chinese strings for chart elements. | Show how to assign an instance of ChartChineseSettings to Workbook.Settings.GlobalizationSettings.ChartSettings and verify the chart’s display unit label changes from English to Chinese. | Explain the steps to extend ChartGlobalizationSettings for another language, listing the methods that need overriding and how to apply the new settings to a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartChineseDemo
{
    // Subclass of ChartGlobalizationSettings providing Chinese translations
    // Demonstrates how to subclass ChartGlobalizationSettings, override methods to return Chinese strings for axis titles, unit names, legend and series labels, apply the settings to a workbook, and see the chart labels switch from English to Chinese before saving the file.
    public class ChartChineseSettings : ChartGlobalizationSettings
    {
        // Axis title name (e.g., "轴标题")
        public override string GetAxisTitleName()
        {
            return "轴标题";
        }

        // Axis unit name based on display unit type
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

        // Chart title name
        public override string GetChartTitleName()
        {
            return "图表标题";
        }

        // Legend decrease label
        public override string GetLegendDecreaseName()
        {
            return "递减";
        }

        // Legend increase label
        public override string GetLegendIncreaseName()
        {
            return "递增";
        }

        // Legend total label
        public override string GetLegendTotalName()
        {
            return "总计";
        }

        // "Other" label
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

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("类别");
            sheet.Cells["A2"].PutValue("第一季度");
            sheet.Cells["A3"].PutValue("第二季度");
            sheet.Cells["A4"].PutValue("第三季度");
            sheet.Cells["B1"].PutValue("数值");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(250);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "销售数据";

            // Set display unit to Hundreds and enable the unit label
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Hundreds;
            chart.ValueAxis.IsDisplayUnitLabelShown = true;

            // Show default (English) unit label before applying Chinese settings
            Console.WriteLine("Default display unit label: " + chart.ValueAxis.DisplayUnitLabel.Text);

            // Apply custom Chinese globalization settings
            wb.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChartChineseSettings()
            };

            // After applying settings, the unit label should be in Chinese
            Console.WriteLine("Chinese display unit label: " + chart.ValueAxis.DisplayUnitLabel.Text);

            // Change display unit to Thousands to demonstrate different translation
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
            Console.WriteLine("Updated Chinese display unit label: " + chart.ValueAxis.DisplayUnitLabel.Text);

            // Save the workbook (output file)
            wb.Save("ChartChineseSettingsDemo.xlsx");
        }
    }
}
