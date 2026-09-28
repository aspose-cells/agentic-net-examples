// Title: Create a C# unit test to verify that an Aspose.Cells column chart displays the correct Chinese title and legend text
// AI Prompts: Generate an MSTest method that builds a workbook, adds a column chart, sets chart.Title.Text to a Chinese string, assigns a Chinese series name, and uses Assert.AreEqual to validate both properties. | Write a NUnit test case that creates a chart with Aspose.Cells, applies Chinese characters to the chart title and series name, then asserts the values match the expected strings. | Provide an xUnit test that constructs a workbook, inserts a column chart, sets Chinese text for the title and legend, and checks the properties with FluentAssertions.
// Common Searches: how to assert Chinese characters in Aspose.Cells chart title using MSTest | unit testing Aspose.Cells chart legend localization in C# | verify column chart title text is Chinese with Aspose.Cells .NET | C# test for Aspose.Cells chart series name containing Chinese characters
// Tags: Aspose.Cells unit test chart title | C# verify chart legend Chinese text | Aspose.Cells column chart localization | assert chart.Title.Text Aspose.Cells | test Aspose.Cells NSeries name

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for ChartType enum

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add a column chart, assign Chinese strings to the chart title and series name, and write a unit test that asserts those properties using popular .NET testing frameworks.
    public class ChartChineseSettingsDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Populate data for the chart series
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);
                sheet.Cells["B1"].PutValue(15);
                sheet.Cells["B2"].PutValue(25);
                sheet.Cells["B3"].PutValue(35);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
                var chart = sheet.Charts[chartIndex];

                // Set Chinese title for the chart
                string chineseTitle = "销售额统计";
                chart.Title.Text = chineseTitle;
                chart.Title.IsVisible = true;

                // Add a series and set a Chinese name (appears in the legend)
                chart.NSeries.Add("A1:A3", true);
                chart.NSeries[0].Name = "第一季度";

                // Retrieve the title and legend name
                string actualTitle = chart.Title.Text;
                string actualLegend = chart.NSeries[0].Name;

                // Verify that the title and legend contain the expected Chinese strings
                if (actualTitle == chineseTitle && actualLegend == "第一季度")
                {
                    Console.WriteLine("Test passed: Title and legend are correctly set to Chinese strings.");
                }
                else
                {
                    Console.WriteLine("Test failed:");
                    Console.WriteLine($"Expected title: {chineseTitle}, Actual title: {actualTitle}");
                    Console.WriteLine($"Expected legend: 第一季度, Actual legend: {actualLegend}");
                }

                // Optional: save the workbook to inspect the chart manually
                // workbook.Save("ChartChineseSettings.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
