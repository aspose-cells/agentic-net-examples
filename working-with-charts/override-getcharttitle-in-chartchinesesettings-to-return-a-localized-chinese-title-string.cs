// Title: C# – Override SettableChartGlobalizationSettings to add a Chinese chart title in Aspose.Cells
// Description: Demonstrates how to subclass SettableChartGlobalizationSettings, override GetChartTitleName to return "图表标题", and apply the custom title to a column chart in a workbook created with Aspose.Cells for .NET. The example populates sample data, creates the chart, sets the localized title, makes it visible, and saves the file as ChartWithChineseTitle.xlsx.
// Keywords: Aspose.Cells | C# | SettableChartGlobalizationSettings | GetChartTitleName | Chinese chart title | chart globalization | Excel localization | China | chart title override | Aspose.Cells example
// Common Searches: Aspose.Cells set Chinese chart title | override GetChartTitleName C# | custom chart globalization Aspose.Cells | how to localize Excel chart titles in .NET | C# Aspose.Cells Chinese localization example
// Developer Intent: Show how to provide a Chinese‑language title for an Excel chart by overriding the globalization settings in Aspose.Cells for .NET.
// Use Cases: Create Excel reports for Chinese‑speaking users with automatically localized chart titles. | Reuse a ChartChineseSettings class across projects to enforce consistent Chinese chart labeling. | Combine the custom title with additional formatting (axis labels, legends) for multilingual dashboards.
// AI Prompts: Write C# code that defines a SettableChartGlobalizationSettings subclass returning the Chinese string "图表标题" and applies it to a chart in Aspose.Cells. | Explain step‑by‑step how to override GetChartTitleName to localize a chart title and make it visible in a column chart. | Give examples of extending ChartChineseSettings to also translate axis titles and legend text in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartChineseDemo
{
    // Custom globalization settings that provide a Chinese chart title
    // Demonstrates how to subclass SettableChartGlobalizationSettings, override GetChartTitleName to return "图表标题", and apply the custom title to a column chart in a workbook created with Aspose.Cells for .NET. The example populates sample data, creates the chart, sets the localized title, makes it visible, and saves the file as ChartWithChineseTitle.xlsx.
    public class ChartChineseSettings : SettableChartGlobalizationSettings
    {
        // Override the method to return a localized Chinese title
        public override string GetChartTitleName()
        {
            // "图表标题" means "Chart Title" in Chinese
            return "图表标题";
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            worksheet.Cells["A1"].PutValue("类别");
            worksheet.Cells["A2"].PutValue("第一季度");
            worksheet.Cells["A3"].PutValue("第二季度");
            worksheet.Cells["B1"].PutValue("数值");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(150);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Use the custom Chinese globalization settings to obtain the title
            ChartChineseSettings chineseSettings = new ChartChineseSettings();
            chart.Title.Text = chineseSettings.GetChartTitleName();
            chart.Title.IsVisible = true;

            // Save the workbook
            workbook.Save("ChartWithChineseTitle.xlsx");
        }
    }
}
