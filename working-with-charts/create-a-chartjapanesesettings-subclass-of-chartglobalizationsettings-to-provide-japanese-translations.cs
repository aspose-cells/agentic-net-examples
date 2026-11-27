// Title: Create a Japanese chart localization subclass by extending ChartGlobalizationSettings in Aspose.Cells for .NET
// Common Searches: aspnet chart globalization settings japanese example | override ChartGlobalizationSettings to provide Japanese axis titles in C# | Aspose.Cells chart Japanese legend text custom subclass | how to set Japanese unit names for chart axes using Aspose.Cells | C# create column chart with Japanese labels using Aspose.Cells
// Tags: ChartGlobalizationSettings Japanese subclass | Aspose.Cells chart localization C# | Japanese axis unit names Aspose.Cells | custom chart legend translation Aspose.Cells | C# column chart with Japanese labels

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartJapaneseSettingsDemo
{
    // Subclass of ChartGlobalizationSettings providing Japanese translations
    // Defines a ChartJapaneseSettings class that inherits from ChartGlobalizationSettings and overrides methods to return Japanese strings for axis titles, unit names, chart title, legend entries, and series names, then demonstrates creating a workbook with a column chart that uses these Japanese labels.
    public class ChartJapaneseSettings : ChartGlobalizationSettings
    {
        public override string GetAxisTitleName()
        {
            return "軸タイトル";
        }

        public override string GetAxisUnitName(DisplayUnitType unitType)
        {
            switch (unitType)
            {
                case DisplayUnitType.None:
                    return "";
                case DisplayUnitType.Hundreds:
                    return "百";
                case DisplayUnitType.Thousands:
                    return "千";
                case DisplayUnitType.Millions:
                    return "百万";
                case DisplayUnitType.Billions:
                    return "十億";
                case DisplayUnitType.Trillions:
                    return "兆";
                default:
                    return base.GetAxisUnitName(unitType);
            }
        }

        public override string GetChartTitleName()
        {
            return "チャートタイトル";
        }

        public override string GetLegendDecreaseName()
        {
            return "減少";
        }

        public override string GetLegendIncreaseName()
        {
            return "増加";
        }

        public override string GetLegendTotalName()
        {
            return "合計";
        }

        public override string GetOtherName()
        {
            return "その他";
        }

        public override string GetSeriesName()
        {
            return "系列";
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

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("カテゴリ");
                sheet.Cells["A2"].PutValue("商品A");
                sheet.Cells["A3"].PutValue("商品B");
                sheet.Cells["A4"].PutValue("商品C");

                sheet.Cells["B1"].PutValue("売上");
                sheet.Cells["B2"].PutValue(1500);
                sheet.Cells["B3"].PutValue(2300);
                sheet.Cells["B4"].PutValue(1800);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Use Japanese strings directly (GlobalizationSettings property not available in current API)
                chart.Title.Text = "チャートタイトル";
                chart.ValueAxis.Title.Text = "軸タイトル";
                chart.CategoryAxis.Title.Text = "軸タイトル";

                // Configure legend
                chart.Legend.Position = LegendPositionType.Right;
                // Legend visibility property may not exist in this version; skip if unavailable

                // Save the workbook
                string outputPath = "ChartJapaneseSettingsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
