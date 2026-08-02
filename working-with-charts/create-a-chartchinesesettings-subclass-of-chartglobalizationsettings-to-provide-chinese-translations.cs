// Title: Localize Aspose.Cells Chart Labels to Chinese with a ChartGlobalizationSettings Subclass (C#)
// Description: This example shows how to create a ChartChineseSettings class that inherits from ChartGlobalizationSettings and overrides methods to supply Chinese strings for axis titles, display units, chart title, legend entries, "Other" and series names. The custom settings are applied to a workbook via GlobalizationSettings.ChartSettings, a column chart is generated, and the file is saved as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | ChartGlobalizationSettings | Chinese localization | Excel chart labels | chart axis Chinese | legend translation | globalization settings | multilingual Excel | custom chart settings
// Common Searches: Aspose.Cells chart Chinese localization example | how to set Chinese axis titles in Aspose.Cells | custom ChartGlobalizationSettings for Chinese labels | apply Chinese legend text to Aspose.Cells charts | C# chart globalization settings Aspose.Cells
// Developer Intent: Implement a subclass of ChartGlobalizationSettings that returns Chinese text for all chart UI elements and apply it to a workbook to produce fully localized charts.
// Use Cases: Produce Excel reports for Chinese audiences with automatically translated chart titles, axes, and legends. | Standardize Chinese chart terminology across multiple workbooks in a multilingual .NET application. | Create reusable localization components for Aspose.Cells charts without manually editing each label.
// AI Prompts: Generate a ChartGlobalizationSettings subclass that provides Japanese translations for chart elements in Aspose.Cells. | Explain how to switch between English and Chinese ChartGlobalizationSettings at runtime in a C# workbook. | Show code to localize data label text for French using a custom Aspose.Cells globalization class.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDemo
{
    // Custom Chinese globalization settings for charts
    // This example shows how to create a ChartChineseSettings class that inherits from ChartGlobalizationSettings and overrides methods to supply Chinese strings for axis titles, display units, chart title, legend entries, "Other" and series names. The custom settings are applied to a workbook via GlobalizationSettings.ChartSettings, a column chart is generated, and the file is saved as an Excel workbook.
    public class ChartChineseSettings : ChartGlobalizationSettings
    {
        // Axis title name in Chinese
        public override string GetAxisTitleName()
        {
            return "坐标轴标题";
        }

        // Axis unit names for different display unit types
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

        // Chart title name in Chinese
        public override string GetChartTitleName()
        {
            return "图表标题";
        }

        // Legend decrease label in Chinese
        public override string GetLegendDecreaseName()
        {
            return "递减";
        }

        // Legend increase label in Chinese
        public override string GetLegendIncreaseName()
        {
            return "递增";
        }

        // Legend total label in Chinese
        public override string GetLegendTotalName()
        {
            return "合计";
        }

        // "Other" label in Chinese
        public override string GetOtherName()
        {
            return "其他";
        }

        // Series name in Chinese
        public override string GetSeriesName()
        {
            return "系列";
        }
    }

    // Demonstration of applying the custom Chinese settings to a workbook
    public static class Demo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("类别");
            ws.Cells["A2"].PutValue("第一类");
            ws.Cells["A3"].PutValue("第二类");
            ws.Cells["B1"].PutValue("数值");
            ws.Cells["B2"].PutValue(120);
            ws.Cells["B3"].PutValue(340);

            // Add a column chart
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";
            chart.Title.Text = "示例图表";

            // Apply the custom Chinese globalization settings
            wb.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChartChineseSettings()
            };

            // Save the workbook
            wb.Save("ChartChineseDemo.xlsx");
        }
    }

    // Program entry point
    public class Program
    {
        public static void Main()
        {
            try
            {
                Demo.Run();
                Console.WriteLine("Workbook created successfully: ChartChineseDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
