// Title: Localize Aspose.Cells Chart Title to Chinese with ChartChineseSettings (.NET)
// Description: This example shows how to create a ChartChineseSettings class derived from SettableChartGlobalizationSettings, override its GetChartTitleName method to return “图表标题”, set the title on a column chart, make it visible, and save the workbook as an .xlsx file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | chart title localization | Chinese title | SettableChartGlobalizationSettings | GetChartTitleName | .NET | C# | Excel chart customization | globalization settings | custom chart title
// Common Searches: Aspose.Cells Chinese chart title | override GetChartTitleName C# example | custom ChartGlobalizationSettings Aspose.Cells | localize Excel chart titles with .NET | set chart title text programmatically Aspose.Cells
// Developer Intent: Create and use a custom globalization settings class that supplies a Chinese title for Excel charts generated with Aspose.Cells.
// Use Cases: Produce Chinese‑language financial reports with correctly localized chart headings. | Apply the same ChartChineseSettings across multiple worksheets to maintain consistent terminology. | Package the custom settings in a shared NuGet library for reuse in enterprise reporting solutions.
// AI Prompts: Generate C# code that derives from SettableChartGlobalizationSettings to return a Japanese chart title and applies it to an Aspose.Cells chart. | Describe the steps to integrate a custom ChartGlobalizationSettings implementation into an existing Aspose.Cells project for multi‑culture chart title support. | Provide a verification checklist to confirm that the overridden GetChartTitleName method updates the chart title in the saved Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Custom globalization settings that provide a Chinese chart title
    // This example shows how to create a ChartChineseSettings class derived from SettableChartGlobalizationSettings, override its GetChartTitleName method to return “图表标题”, set the title on a column chart, make it visible, and save the workbook as an .xlsx file using Aspose.Cells for .NET.
    public class ChartChineseSettings : SettableChartGlobalizationSettings
    {
        // Override the base method to return a localized Chinese title
        public override string GetChartTitleName()
        {
            // "图表标题" means "Chart Title" in Chinese
            return "图表标题";
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data for the chart
                sheet.Cells["A1"].PutValue("类别");
                sheet.Cells["A2"].PutValue("第一");
                sheet.Cells["A3"].PutValue("第二");
                sheet.Cells["B1"].PutValue("值");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Use the custom settings to obtain the Chinese title and assign it to the chart
                var chineseSettings = new ChartChineseSettings();
                chart.Title.Text = chineseSettings.GetChartTitleName();
                chart.Title.IsVisible = true;

                // Save the workbook to a file
                string outputPath = "ChartChineseSettingsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
