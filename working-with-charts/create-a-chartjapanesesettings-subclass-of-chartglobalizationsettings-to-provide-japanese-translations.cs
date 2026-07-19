// Title: C# – Localize Aspose.Cells Chart Labels to Japanese with a ChartJapaneseSettings Subclass
// Description: This example shows how to subclass ChartGlobalizationSettings in Aspose.Cells to supply Japanese text for axis titles, units, chart title, legend entries, and series names. The custom ChartJapaneseSettings object is assigned to Workbook.Settings.GlobalizationSettings.ChartSettings, producing a column chart with fully Japanese UI that is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | ChartGlobalizationSettings | Japanese localization | ChartJapaneseSettings | GlobalizationSettings.ChartSettings | Excel chart Japanese labels | .NET chart globalization | DisplayUnitType Japanese | column chart example | localize chart axis | Japanese legend text
// Common Searches: Aspose.Cells Japanese chart localization | C# ChartGlobalizationSettings Japanese example | how to set chart language to Japanese in Aspose.Cells | override GetAxisUnitName for Japanese units | GlobalizationSettings.ChartSettings Aspose.Cells tutorial | Japanese labels on Excel charts using Aspose | ChartJapaneseSettings class sample
// Developer Intent: Create a custom ChartGlobalizationSettings subclass that returns Japanese strings for every chart UI element and apply it through GlobalizationSettings to generate Japanese‑language Excel charts.
// Use Cases: Produce Excel reports with charts that display Japanese axis titles, units, and legends for a Japanese‑speaking audience. | Integrate chart localization into automated .NET reporting pipelines that use Aspose.Cells. | Customize financial axis units (千, 万, 百万, 十億) for Japanese market data visualizations. | Reuse the ChartJapaneseSettings class across multiple workbooks to ensure consistent Japanese chart terminology. | Localize legend entries such as 増加, 減少, 合計, その他 for column, line, or pie charts.
// AI Prompts: Generate a C# class inheriting from ChartGlobalizationSettings that returns Japanese text for axis titles, units, chart title, legend labels, and series names. | Demonstrate how to assign a ChartJapaneseSettings instance to workbook.Settings.GlobalizationSettings.ChartSettings and create a column chart with Japanese UI. | List all overridable methods of ChartGlobalizationSettings required to fully translate a chart to Japanese. | Explain the mapping between DisplayUnitType values and Japanese unit names in ChartJapaneseSettings. | Provide a step‑by‑step tutorial for localizing Aspose.Cells charts to Japanese, including code snippets and saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsJapaneseChartDemo
{
    // Subclass of ChartGlobalizationSettings providing Japanese translations
    // This example shows how to subclass ChartGlobalizationSettings in Aspose.Cells to supply Japanese text for axis titles, units, chart title, legend entries, and series names. The custom ChartJapaneseSettings object is assigned to Workbook.Settings.GlobalizationSettings.ChartSettings, producing a column chart with fully Japanese UI that is saved as an XLSX file.
    public class ChartJapaneseSettings : ChartGlobalizationSettings
    {
        // Axis title (e.g., "軸タイトル")
        public override string GetAxisTitleName()
        {
            return "軸タイトル";
        }

        // Axis unit based on display unit type
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
                    return "十億";
                default:
                    return base.GetAxisUnitName(type);
            }
        }

        // Chart title (e.g., "チャートタイトル")
        public override string GetChartTitleName()
        {
            return "チャートタイトル";
        }

        // Legend decrease (e.g., "減少")
        public override string GetLegendDecreaseName()
        {
            return "減少";
        }

        // Legend increase (e.g., "増加")
        public override string GetLegendIncreaseName()
        {
            return "増加";
        }

        // Legend total (e.g., "合計")
        public override string GetLegendTotalName()
        {
            return "合計";
        }

        // "Other" label (e.g., "その他")
        public override string GetOtherName()
        {
            return "その他";
        }

        // Series name (e.g., "系列")
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
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("カテゴリ");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("値");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(250);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply Japanese chart globalization settings
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChartJapaneseSettings()
            };

            // Set chart title to demonstrate GetChartTitleName usage (optional)
            chart.Title.Text = "売上チャート";

            // Save the workbook
            workbook.Save("JapaneseChartDemo.xlsx");
        }
    }
}
