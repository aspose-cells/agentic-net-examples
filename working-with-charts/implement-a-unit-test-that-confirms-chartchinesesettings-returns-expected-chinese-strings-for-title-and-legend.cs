// Title: Unit test for Chinese chart globalization settings in Aspose.Cells for .NET
// Description: Demonstrates how to create a custom ChineseChartGlobalizationSettings class that overrides ChartGlobalizationSettings, apply it to a Workbook, assert that GetChartTitleName returns "图表标题" and GetLegendIncreaseName returns "增加", and verify the workbook saves without errors.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | Chinese localization | C# unit test | MSTest | xUnit | NUnit | .NET | chart title | legend text | globalization settings
// Common Searches: Aspose.Cells test custom chart globalization | verify Chinese chart title Aspose.Cells | unit test ChartGlobalizationSettings C# | how to assert chart legend text in Aspose.Cells | sample code for chart localization Aspose.Cells .NET
// Developer Intent: Write a unit test that confirms the overridden Chinese strings for chart title and legend are returned correctly by Aspose.Cells.
// Use Cases: Automatically validate that custom Chinese chart globalization settings are applied during workbook generation. | Prevent regressions when upgrading Aspose.Cells by testing localization overrides. | Ensure that saving a workbook after setting ChineseChartGlobalizationSettings does not modify the overridden values.
// AI Prompts: Generate an MSTest method that creates a Workbook, sets GlobalizationSettings.ChartSettings to ChineseChartGlobalizationSettings, and asserts GetChartTitleName == "图表标题" and GetLegendIncreaseName == "增加". | Provide an xUnit test example for verifying custom ChartGlobalizationSettings in Aspose.Cells, including proper disposal of the Workbook and cleanup of temporary files. | Write a NUnit test case that checks Chinese chart globalization overrides and confirms the workbook saves without throwing exceptions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTests
{
    // Custom Chinese globalization settings for charts
    // Demonstrates how to create a custom ChineseChartGlobalizationSettings class that overrides ChartGlobalizationSettings, apply it to a Workbook, assert that GetChartTitleName returns "图表标题" and GetLegendIncreaseName returns "增加", and verify the workbook saves without errors.
    public class ChineseChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // Returns Chinese string for chart title
        public override string GetChartTitleName()
        {
            return "图表标题";
        }

        // Returns Chinese string for legend increase
        public override string GetLegendIncreaseName()
        {
            return "增加";
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Apply the custom Chinese chart globalization settings
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new ChineseChartGlobalizationSettings()
                };

                // Retrieve the settings via the workbook
                ChartGlobalizationSettings settings = workbook.Settings.GlobalizationSettings.ChartSettings;

                // Verify that the Chinese strings are returned as expected
                if (settings.GetChartTitleName() != "图表标题")
                {
                    throw new InvalidOperationException("Chart title name does not match expected Chinese value.");
                }

                if (settings.GetLegendIncreaseName() != "增加")
                {
                    throw new InvalidOperationException("Legend increase name does not match expected Chinese value.");
                }

                // Save the workbook to ensure the lifecycle save rule is exercised
                string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                workbook.Save(tempFile);
                Console.WriteLine($"Workbook saved successfully to: {tempFile}");
            }
            catch (Exception ex)
            {
                // Runtime safety: capture any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
