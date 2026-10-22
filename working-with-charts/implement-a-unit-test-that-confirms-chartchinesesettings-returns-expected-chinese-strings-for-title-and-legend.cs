// Title: Write a C# unit test using Aspose.Cells to validate that a chart’s Title.Text and Legend.Text contain the expected Chinese strings
// AI Prompts: Generate an MSTest method that creates a Workbook, adds a column chart, sets chart.Title.Text to "销售额统计" and chart.Legend.Text to "月份", then asserts both properties match the expected Chinese values. | Provide a NUnit test case for Aspose.Cells that verifies the Chinese title and legend of a column chart are stored and retrieved correctly.
// Common Searches: aspocells c# unit test verify chart title Chinese characters | how to assert Chinese legend text in Aspose.Cells chart unit test | C# Aspose.Cells column chart localization test example | testing chart Title.Text with Chinese strings using Aspose.Cells | unit testing chart Chinese labels in Aspose.Cells workbook
// Tags: Aspose.Cells chart title verification | C# unit test Aspose.Cells chart localization | column chart Chinese text assertion | Aspose.Cells Title.Text Chinese validation | Aspose.Cells Legend.Text unit test

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTests
{
    // The example creates a Workbook, fills cells with Chinese category labels and numeric data, adds a column chart, sets the chart's Title.Text to "销售额统计" and Legend.Text to "月份", then retrieves these properties and asserts they equal the expected Chinese strings, returning true only when both match.
    public class ChartChineseSettingsTests
    {
        public static void Main()
        {
            try
            {
                bool result = RunChartChineseSettingsTest();
                Console.WriteLine(result
                    ? "ChartChineseSettings test passed."
                    : "ChartChineseSettings test failed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        private static bool RunChartChineseSettingsTest()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Add data for the chart
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("一");
                sheet.Cells["A3"].PutValue("二");
                sheet.Cells["A4"].PutValue("三");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                var chart = sheet.Charts[chartIndex];

                // Set data source
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply Chinese settings using Title and Legend text properties
                chart.Title.Text = "销售额统计";
                chart.Legend.Text = "月份";

                // Retrieve settings
                string actualTitle = chart.Title.Text;
                string actualLegend = chart.Legend.Text;

                // Verify expected values
                bool titleMatches = actualTitle == "销售额统计";
                bool legendMatches = actualLegend == "月份";

                return titleMatches && legendMatches;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test execution error: {ex.Message}");
                return false;
            }
        }
    }
}
