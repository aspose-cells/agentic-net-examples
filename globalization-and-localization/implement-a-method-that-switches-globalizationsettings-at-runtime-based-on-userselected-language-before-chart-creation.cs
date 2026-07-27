using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartGlobalizationDemo
{
    // Helper class to configure chart globalization settings based on language
    public static class ChartGlobalizationHelper
    {
        // Creates a workbook, applies language‑specific chart globalization, adds data and a chart
        public static Workbook CreateChartWithLanguage(string language)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B1"].PutValue("Series");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Create language‑specific chart globalization settings
            SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();

            switch (language?.ToLowerInvariant())
            {
                case "fr": // French
                    chartSettings.SetChartTitleName("Titre du graphique");
                    chartSettings.SetSeriesName("Série");
                    chartSettings.SetLegendIncreaseName("Augmenter");
                    chartSettings.SetLegendDecreaseName("Diminuer");
                    chartSettings.SetOtherName("Autre");
                    break;

                case "zh": // Chinese
                    chartSettings.SetChartTitleName("图表标题");
                    chartSettings.SetSeriesName("系列");
                    chartSettings.SetLegendIncreaseName("增加");
                    chartSettings.SetLegendDecreaseName("减少");
                    chartSettings.SetOtherName("其他");
                    break;

                default: // Default (English) – no custom text needed
                    // Optionally you could set English explicitly
                    chartSettings.SetChartTitleName("Chart Title");
                    chartSettings.SetSeriesName("Series");
                    chartSettings.SetLegendIncreaseName("Increase");
                    chartSettings.SetLegendDecreaseName("Decrease");
                    chartSettings.SetOtherName("Other");
                    break;
            }

            // Apply the chart globalization settings to the workbook
            GlobalizationSettings globalization = new GlobalizationSettings
            {
                ChartSettings = chartSettings
            };
            workbook.Settings.GlobalizationSettings = globalization;

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // The chart will automatically use the localized strings set above
            return workbook;
        }
    }

    // Demo program
    class Program
    {
        static void Main()
        {
            // Example: create charts for English, French and Chinese
            Workbook wbEn = ChartGlobalizationHelper.CreateChartWithLanguage("en");
            wbEn.Save("Chart_English.xlsx");

            Workbook wbFr = ChartGlobalizationHelper.CreateChartWithLanguage("fr");
            wbFr.Save("Chart_French.xlsx");

            Workbook wbZh = ChartGlobalizationHelper.CreateChartWithLanguage("zh");
            wbZh.Save("Chart_Chinese.xlsx");

            Console.WriteLine("Workbooks saved with language‑specific chart globalization.");
        }
    }
}