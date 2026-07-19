// Title: Set Japanese Globalization Settings for an Aspose.Cells Chart in C#
// Description: Shows how to create a workbook, add sample data and a column chart, define a custom ChartJapaneseSettings class that localizes chart titles, axis titles, legend entries, and series names to Japanese, assign it to the workbook's GlobalizationSettings, and save the file so the exported chart displays Japanese labels.
// Keywords: Aspose.Cells | C# chart globalization | Japanese chart localization | ChartJapaneseSettings | globalization settings Aspose.Cells | export chart Japanese labels | Aspose.Cells chart localization | SetChartTitleName | SetAxisTitleName | SetLegendIncreaseName | SetLegendDecreaseName | SetSeriesName | Excel export Japanese
// Common Searches: Aspose.Cells set Japanese labels for chart elements | How to apply custom ChartJapaneseSettings in .NET | GlobalizationSettings for charts in Aspose.Cells | Export Excel workbook with Japanese chart titles | C# example for chart localization with Aspose.Cells
// Developer Intent: Assign a ChartJapaneseSettings instance to the workbook’s GlobalizationSettings so the exported chart uses Japanese UI strings.
// Use Cases: Produce sales charts for Japanese market reports with fully localized axis and legend text. | Create Excel files that comply with regional language requirements by localizing all chart UI strings to Japanese. | Reuse a single ChartJapaneseSettings object across multiple charts to maintain consistent Japanese terminology throughout a workbook.
// AI Prompts: Generate C# code that applies a custom ChartJapaneseSettings object to a chart’s GlobalizationSettings in Aspose.Cells before saving. | Explain the difference between setting workbook.Settings.GlobalizationSettings.ChartSettings and chart.GlobalizationSettings in Aspose.Cells. | Provide an example of a line chart with Japanese titles, axis labels, and legend entries using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartJapaneseSettingsDemo
{
    // Custom globalization settings for charts with Japanese labels
    // Shows how to create a workbook, add sample data and a column chart, define a custom ChartJapaneseSettings class that localizes chart titles, axis titles, legend entries, and series names to Japanese, assign it to the workbook's GlobalizationSettings, and save the file so the exported chart displays Japanese labels.
    public class ChartJapaneseSettings : SettableChartGlobalizationSettings
    {
        public ChartJapaneseSettings()
        {
            // Set Japanese names for various chart elements
            SetChartTitleName("チャートタイトル");
            SetAxisTitleName("軸タイトル");
            SetLegendIncreaseName("増加");
            SetLegendDecreaseName("減少");
            SetOtherName("その他");
            SetSeriesName("シリーズ");
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
            sheet.Cells["B1"].PutValue("値");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "売上チャート";

            // Apply Japanese globalization settings to the chart
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChartJapaneseSettings()
            };

            // Save the workbook (export)
            workbook.Save("ChartWithJapaneseSettings.xlsx");
        }
    }
}
