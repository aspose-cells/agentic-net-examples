// Title: How to subclass ChartGlobalizationSettings to provide Chinese translations for Aspose.Cells charts in C#
// Common Searches: C# subclass ChartGlobalizationSettings to display Chinese axis titles in Aspose.Cells | Aspose.Cells chart localization to Chinese using custom globalization settings | override GetAxisUnitName for Chinese units in an Aspose.Cells chart | set Chinese legend text and chart title in Aspose.Cells with C# | create a column chart with Chinese labels using Aspose.Cells for .NET | apply Chinese translations to Aspose.Cells chart elements programmatically
// Tags: Aspose.Cells chart globalization subclass C# | Chinese chart axis title translation Aspose.Cells | override GetAxisUnitName for Chinese units | Aspose.Cells legend label localization Chinese | C# column chart with Chinese labels Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartChineseLocalizationDemo
{
    // Subclass of ChartGlobalizationSettings providing Chinese translations
    // The example defines a ChartChineseSettings class that inherits from ChartGlobalizationSettings and overrides methods to return Chinese strings for chart titles, axis titles, series names, legend labels, and display unit names. A workbook is created, sample data is added, a column chart is generated, and the custom Chinese settings are applied to the chart's title, axes, and legend before saving the file as an .xlsx workbook.
    public class ChartChineseSettings : ChartGlobalizationSettings
    {
        // Axis title (e.g., "轴标题")
        public override string GetAxisTitleName()
        {
            return "轴标题";
        }

        // Chart title (e.g., "图表标题")
        public override string GetChartTitleName()
        {
            return "图表标题";
        }

        // Series name (e.g., "系列")
        public override string GetSeriesName()
        {
            return "系列";
        }

        // "Other" label (e.g., "其他")
        public override string GetOtherName()
        {
            return "其他";
        }

        // Legend increase (e.g., "增加")
        public override string GetLegendIncreaseName()
        {
            return "增加";
        }

        // Legend decrease (e.g., "减少")
        public override string GetLegendDecreaseName()
        {
            return "减少";
        }

        // Legend total (e.g., "总计")
        public override string GetLegendTotalName()
        {
            return "总计";
        }

        // Axis unit name based on display unit type
        public override string GetAxisUnitName(DisplayUnitType unitType)
        {
            // Simple mapping; extend as needed
            return unitType switch
            {
                DisplayUnitType.None => "",
                DisplayUnitType.Hundreds => "百",
                DisplayUnitType.Thousands => "千",
                DisplayUnitType.Millions => "百万",
                DisplayUnitType.Billions => "十亿",
                DisplayUnitType.Trillions => "万亿",
                _ => base.GetAxisUnitName(unitType),
            };
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("类别");
                sheet.Cells["A2"].PutValue("苹果");
                sheet.Cells["A3"].PutValue("橙子");
                sheet.Cells["A4"].PutValue("香蕉");

                sheet.Cells["B1"].PutValue("销量");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["B4"].PutValue(1500);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Instantiate Chinese globalization settings
                ChartChineseSettings chineseSettings = new ChartChineseSettings();

                // Apply translated titles/labels
                chart.Title.Text = chineseSettings.GetChartTitleName();
                chart.ValueAxis.Title.Text = chineseSettings.GetAxisTitleName();
                chart.CategoryAxis.Title.Text = chineseSettings.GetAxisTitleName();

                // Example of using legend names
                chart.Legend.Position = LegendPositionType.Right;
                chart.Legend.Text = chineseSettings.GetLegendTotalName();

                // Save the workbook
                string outputPath = "ChartChineseLocalizationDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
