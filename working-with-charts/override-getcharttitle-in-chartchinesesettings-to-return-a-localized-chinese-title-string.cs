// Title: Override GetChartTitleName in a SettableChartGlobalizationSettings subclass to apply a Chinese chart title with Aspose.Cells for .NET
// AI Prompts: Create a C# subclass of SettableChartGlobalizationSettings that overrides GetChartTitleName to return a Chinese string and use it to assign a chart title in Aspose.Cells. | Show how to set the Title.Text property of a column chart to the value returned by the overridden GetChartTitleName and enable the title's visibility before saving the workbook. | Provide a full Aspose.Cells example that populates worksheet data, adds a column chart, applies the custom Chinese title via the overridden method, and saves the file as an .xlsx.
// Common Searches: aspnet override GetChartTitleName for Chinese chart title in Aspose.Cells | example of custom SettableChartGlobalizationSettings to localize chart titles in .NET | how to set chart title text programmatically using Aspose.Cells C#
// Tags: override GetChartTitleName Aspose.Cells | custom SettableChartGlobalizationSettings implementation | set chart title text C# | Chinese localization for chart titles | save workbook with localized chart title .xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Custom class that provides a Chinese chart title
    // This example defines a ChartChineseSettings class that inherits SettableChartGlobalizationSettings and overrides GetChartTitleName to return the Chinese string "图表标题". It creates a workbook, adds sample data and a column chart, sets the chart's Title.Text to the overridden value, makes the title visible, and saves the workbook as ChartWithChineseTitle.xlsx.
    public class ChartChineseSettings : SettableChartGlobalizationSettings
    {
        // Return a Chinese localized title
        public override string GetChartTitleName()
        {
            // "图表标题" means "Chart Title" in Chinese
            return "图表标题";
        }
    }

    public class ChartChineseSettingsDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("类别");
                sheet.Cells["A2"].PutValue("第一类");
                sheet.Cells["A3"].PutValue("第二类");
                sheet.Cells["B1"].PutValue("数值");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(250);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Use the custom Chinese title (globalization settings are not required for this version)
                var chineseSettings = new ChartChineseSettings();
                chart.Title.Text = chineseSettings.GetChartTitleName();
                chart.Title.IsVisible = true;

                // Save the workbook
                workbook.Save("ChartWithChineseTitle.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
