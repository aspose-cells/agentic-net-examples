// Title: C# unit test for Aspose.Cells chart Chinese title and legend text verification
// Common Searches: how to assert Chinese chart title in Aspose.Cells using C# unit test | Aspose.Cells verify chart legend name Chinese characters in C# | unit testing chart localization with Aspose.Cells C# example | C# test for chart.Title.Text Chinese string Aspose.Cells | validate series name Chinese text in Aspose.Cells chart unit test
// Tags: Aspose.Cells chart title verification C# | Aspose.Cells legend localization unit test | C# chart property assertion Aspose.Cells | Chinese string validation Aspose.Cells chart | chart series name check Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTests
{
    // The example creates a workbook, adds a column chart, sets the chart title to the Chinese string "标题", adds a series named "图例", and then uses assertions to confirm that both the Title.Text and the series Name match the expected Chinese values, outputting a success message when the test passes.
    public class ChartChineseSettingsTests
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Get the first worksheet
                var sheet = workbook.Worksheets[0];

                // Add a column chart to the worksheet (Add returns the chart index)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                var chart = sheet.Charts[chartIndex];

                // Expected Chinese strings for title and legend
                const string expectedTitle = "标题";
                const string expectedLegend = "图例";

                // Set the chart title to the expected Chinese text
                chart.Title.Text = expectedTitle;

                // Add a dummy series so that a legend entry exists
                // The series data range can be empty; we only need the legend text
                chart.NSeries.Add("A1:A1", true);
                chart.NSeries[0].Name = expectedLegend;

                // Verify that the title returns the expected value
                if (chart.Title.Text != expectedTitle)
                    throw new Exception($"Title mismatch. Expected: '{expectedTitle}', Actual: '{chart.Title.Text}'");

                // Verify that the legend (series name) returns the expected value
                if (chart.NSeries[0].Name != expectedLegend)
                    throw new Exception($"Legend mismatch. Expected: '{expectedLegend}', Actual: '{chart.NSeries[0].Name}'");

                Console.WriteLine("Chart Chinese settings test passed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
