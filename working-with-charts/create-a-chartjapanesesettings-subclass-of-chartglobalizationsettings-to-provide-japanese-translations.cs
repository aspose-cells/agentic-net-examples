// Title: C# – Localize Aspose.Cells Chart Elements to Japanese with a ChartJapaneseSettings Subclass
// Description: Demonstrates how to create a ChartJapaneseSettings class that inherits from ChartGlobalizationSettings, overrides methods to supply Japanese text for axis titles, units, chart title, legend items, and series names, applies the settings to a workbook, and saves the file as JapaneseChartGlobalization.xlsx.
// Keywords: Aspose.Cells | C# | ChartGlobalizationSettings | Japanese localization | Excel chart translation | chart axis Japanese | legend Japanese | Aspose.Cells .NET example
// Common Searches: Aspose.Cells Japanese chart globalization | How to translate chart axis labels to Japanese in C# | Custom ChartGlobalizationSettings subclass example | Set Japanese legend names in Aspose.Cells | Localize Excel chart titles with Aspose.Cells .NET
// Developer Intent: Implement a reusable ChartGlobalizationSettings subclass that provides Japanese terminology for all chart components and apply it to workbooks generated with Aspose.Cells.
// Use Cases: Produce Excel reports for Japanese users with fully localized chart labels and titles. | Standardize Japanese chart terminology across multiple workbooks in an automated reporting pipeline. | Create multilingual Excel dashboards by swapping ChartGlobalizationSettings instances at runtime.
// AI Prompts: Generate a ChartJapaneseSettings method that returns Japanese text for data labels and error bars. | Show code to switch between English and Japanese ChartGlobalizationSettings based on a user‑selected language. | Provide an example of exporting a workbook using ChartJapaneseSettings to PDF while preserving Japanese characters.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsJapaneseGlobalization
{
    // Subclass of ChartGlobalizationSettings providing Japanese translations
    // Demonstrates how to create a ChartJapaneseSettings class that inherits from ChartGlobalizationSettings, overrides methods to supply Japanese text for axis titles, units, chart title, legend items, and series names, applies the settings to a workbook, and saves the file as JapaneseChartGlobalization.xlsx.
    public class ChartJapaneseSettings : ChartGlobalizationSettings
    {
        // Axis title name (e.g., "軸タイトル")
        public override string GetAxisTitleName()
        {
            return "軸タイトル";
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
                    return "十億";
                default:
                    return base.GetAxisUnitName(type);
            }
        }

        // Chart title name
        public override string GetChartTitleName()
        {
            return "チャートタイトル";
        }

        // Legend decrease name
        public override string GetLegendDecreaseName()
        {
            return "減少";
        }

        // Legend increase name
        public override string GetLegendIncreaseName()
        {
            return "増加";
        }

        // Legend total name
        public override string GetLegendTotalName()
        {
            return "合計";
        }

        // "Other" label name
        public override string GetOtherName()
        {
            return "その他";
        }

        // Series name
        public override string GetSeriesName()
        {
            return "シリーズ";
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
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("売上");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "売上推移";

            // Apply Japanese chart globalization settings
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChartJapaneseSettings()
            };

            // Save the workbook
            workbook.Save("JapaneseChartGlobalization.xlsx");
        }
    }
}
