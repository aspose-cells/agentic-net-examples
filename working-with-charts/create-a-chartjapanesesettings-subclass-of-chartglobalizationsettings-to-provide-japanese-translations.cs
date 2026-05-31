using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Custom globalization settings providing Japanese translations for chart elements
    public class ChartJapaneseSettings : ChartGlobalizationSettings
    {
        // Axis title name in Japanese
        public override string GetAxisTitleName()
        {
            return "軸タイトル";
        }

        // Axis unit names in Japanese based on the display unit type
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

        // Chart title name in Japanese
        public override string GetChartTitleName()
        {
            return "チャートタイトル";
        }

        // Legend decrease label in Japanese
        public override string GetLegendDecreaseName()
        {
            return "減少";
        }

        // Legend increase label in Japanese
        public override string GetLegendIncreaseName()
        {
            return "増加";
        }

        // Legend total label in Japanese
        public override string GetLegendTotalName()
        {
            return "合計";
        }

        // "Other" label in Japanese
        public override string GetOtherName()
        {
            return "その他";
        }

        // Series name in Japanese
        public override string GetSeriesName()
        {
            return "シリーズ";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("カテゴリ");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("値");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set display unit to thousands to demonstrate axis unit translation
            chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
            chart.ValueAxis.IsDisplayUnitLabelShown = true;

            // Apply the custom Japanese globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChartJapaneseSettings()
            };

            // Save the workbook
            workbook.Save("ChartJapaneseSettingsDemo.xlsx");
        }
    }
}