using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Custom globalization settings for Japanese axis titles
    public class ChartJapaneseSettings : SettableChartGlobalizationSettings
    {
        // Override to provide a Japanese axis title.
        // This example returns a generic Japanese title that can be used for both axes.
        public override string GetAxisTitleName()
        {
            // "軸タイトル" means "Axis Title" in Japanese.
            return "軸タイトル";
        }
    }

    public class ChartJapaneseSettingsDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart.
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(150);
            worksheet.Cells["B4"].PutValue(180);

            // Instantiate the custom Japanese settings.
            ChartJapaneseSettings japaneseSettings = new ChartJapaneseSettings();

            // Add a column chart.
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply Japanese axis titles using the overridden method.
            // For demonstration, we use the same title for both axes.
            chart.CategoryAxis.Title.Text = japaneseSettings.GetAxisTitleName(); // X‑axis title
            chart.ValueAxis.Title.Text = japaneseSettings.GetAxisTitleName();    // Y‑axis title

            // Make the titles visible.
            chart.CategoryAxis.Title.IsVisible = true;
            chart.ValueAxis.Title.IsVisible = true;

            // Save the workbook.
            workbook.Save("ChartJapaneseSettingsDemo.xlsx");
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            ChartJapaneseSettingsDemo.Run();
        }
    }
}