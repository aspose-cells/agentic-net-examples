// Title: C# – Localize Aspose.Cells Charts to Japanese with a ChartGlobalizationSettings Subclass
// Description: Demonstrates how to create a ChartJapaneseSettings class that inherits from ChartGlobalizationSettings, overrides methods to supply Japanese text for axis titles, units, chart title, legend labels, "Other" and series names, and applies the custom globalization settings to a workbook. The example builds a simple column chart, sets a Japanese chart title, prints localized values, and saves the file as JapaneseChartDemo.xlsx.
// Keywords: Aspose.Cells | C# chart localization | ChartGlobalizationSettings | Japanese chart labels | Aspose.Cells Japanese example | globalization settings | Excel chart Japanese | DisplayUnitType Japanese | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells chart Japanese localization example | How to override ChartGlobalizationSettings in C# | Set Japanese axis titles in Aspose.Cells charts | Custom chart globalization settings for Japanese language | Apply Japanese legend labels to Aspose.Cells workbook
// Developer Intent: Create a reusable ChartGlobalizationSettings subclass that provides Japanese translations for all chart elements and apply it to an Aspose.Cells workbook.
// Use Cases: Generate sales dashboards with fully Japanese chart UI for regional reports. | Standardize Japanese chart terminology across multiple workbooks in an enterprise solution. | Extend the subclass to support additional Japanese unit strings for financial and scientific charts.
// AI Prompts: Write a C# class that inherits from ChartGlobalizationSettings and returns Japanese strings for axis titles, units, legend labels, and series names. | Show code that assigns a custom ChartGlobalizationSettings subclass to a workbook's GlobalizationSettings and verifies the output in the console. | Explain how to expand ChartJapaneseSettings to include custom number formats and data label translations for Japanese charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsJapaneseChartDemo
{
    // Custom globalization settings for Japanese language
    // Demonstrates how to create a ChartJapaneseSettings class that inherits from ChartGlobalizationSettings, overrides methods to supply Japanese text for axis titles, units, chart title, legend labels, "Other" and series names, and applies the custom globalization settings to a workbook. The example builds a simple column chart, sets a Japanese chart title, prints localized values, and saves the file as JapaneseChartDemo.xlsx.
    public class ChartJapaneseSettings : ChartGlobalizationSettings
    {
        // Axis title name in Japanese
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
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply Japanese globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChartJapaneseSettings()
            };

            // Set chart title to demonstrate that the title text can be localized
            chart.Title.Text = "売上チャート";

            // Display some of the localized names in the console
            Console.WriteLine("Axis Title Name: " + chart.Title.Text);
            Console.WriteLine("Series Name: " + ((ChartJapaneseSettings)workbook.Settings.GlobalizationSettings.ChartSettings).GetSeriesName());
            Console.WriteLine("Legend Increase Name: " + ((ChartJapaneseSettings)workbook.Settings.GlobalizationSettings.ChartSettings).GetLegendIncreaseName());
            Console.WriteLine("Axis Unit (Thousands): " + ((ChartJapaneseSettings)workbook.Settings.GlobalizationSettings.ChartSettings).GetAxisUnitName(DisplayUnitType.Thousands));

            // Save the workbook
            workbook.Save("JapaneseChartDemo.xlsx");
        }
    }
}
