using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Custom globalization settings for charts providing Chinese translations
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
                case DisplayUnitType.HundredThousands:
                    return "十万";
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

        // "Other" label for pie/doughnut charts
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

    public class GlobalizationSettingsChartChineseDemo
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
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "季度销售";

            // Set display unit to thousands to see Chinese unit name
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
            chart.ValueAxis.IsDisplayUnitLabelShown = true;

            // Apply custom Chinese globalization settings for charts
            wb.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChartChineseSettings()
            };

            // Save the workbook (Excel format)
            wb.Save("ChartChineseSettingsDemo.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            GlobalizationSettingsChartChineseDemo.Run();
            Console.WriteLine("Workbook saved with Chinese chart globalization settings.");
        }
    }
}