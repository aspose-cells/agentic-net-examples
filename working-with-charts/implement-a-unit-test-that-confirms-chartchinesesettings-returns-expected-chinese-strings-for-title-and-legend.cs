// Title: C# Unit Test for Custom Chinese ChartGlobalizationSettings in Aspose.Cells
// Description: Demonstrates how to subclass ChartGlobalizationSettings to return Chinese labels, apply the subclass to a Workbook, and verify each overridden method with a simple unit‑test style check before saving and cleaning up a temporary file.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | Chinese localization | C# unit test | .NET | custom chart globalization | chart title Chinese | legend Chinese | axis unit Chinese | DisplayUnitType | MSTest | NUnit | XUnit | globalization settings test
// Common Searches: Aspose.Cells unit test for custom ChartGlobalizationSettings | verify Chinese chart titles and legends in .NET | how to override ChartGlobalizationSettings for Chinese language | C# test for GetAxisUnitName Chinese output | Aspose.Cells chart localization unit test example
// Developer Intent: Write a test that confirms a custom ChartGlobalizationSettings subclass returns the expected Chinese strings for chart titles, legends, series, axis titles, and display units.
// Use Cases: Create a ChineseChartGlobalizationSettings class that supplies localized chart text. | Assign the custom settings to Workbook.Settings.GlobalizationSettings.ChartSettings. | Assert each overridden method returns the correct Chinese label within a test framework. | Save the workbook to a temporary file to ensure the settings are applied during the save lifecycle. | Integrate the verification logic into MSTest, NUnit, or XUnit test suites.
// AI Prompts: Generate an MSTest method that asserts all ChineseChartGlobalizationSettings methods return the expected strings. | Provide an XUnit test that saves and reloads a workbook to verify custom chart globalization persists. | Write a code snippet showing how to loop through all DisplayUnitType values and check GetAxisUnitName returns the correct Chinese unit.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTests
{
    // Custom globalization settings that return Chinese strings
    // Demonstrates how to subclass ChartGlobalizationSettings to return Chinese labels, apply the subclass to a Workbook, and verify each overridden method with a simple unit‑test style check before saving and cleaning up a temporary file.
    public class ChineseChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetChartTitleName() => "图表标题";
        public override string GetLegendIncreaseName() => "增加";
        public override string GetLegendDecreaseName() => "减少";
        public override string GetLegendTotalName() => "总计";
        public override string GetSeriesName() => "系列";
        public override string GetAxisTitleName() => "轴标题";
        public override string GetOtherName() => "其他";

        public override string GetAxisUnitName(DisplayUnitType type)
        {
            return type switch
            {
                DisplayUnitType.Hundreds => "百",
                DisplayUnitType.Thousands => "千",
                DisplayUnitType.TenThousands => "万",
                _ => base.GetAxisUnitName(type),
            };
        }
    }

    public class ChartChineseSettingsDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Apply the custom Chinese globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new ChineseChartGlobalizationSettings()
                };

                // Retrieve the applied chart globalization settings
                ChartGlobalizationSettings settings = workbook.Settings.GlobalizationSettings.ChartSettings;

                // Verify that the overridden methods return the expected Chinese strings
                Verify(settings.GetChartTitleName() == "图表标题", "Chart title name should be Chinese.");
                Verify(settings.GetLegendIncreaseName() == "增加", "Legend increase name should be Chinese.");
                Verify(settings.GetLegendDecreaseName() == "减少", "Legend decrease name should be Chinese.");
                Verify(settings.GetLegendTotalName() == "总计", "Legend total name should be Chinese.");
                Verify(settings.GetSeriesName() == "系列", "Series name should be Chinese.");
                Verify(settings.GetAxisTitleName() == "轴标题", "Axis title name should be Chinese.");
                Verify(settings.GetOtherName() == "其他", "Other name should be Chinese.");
                Verify(settings.GetAxisUnitName(DisplayUnitType.Thousands) == "千", "Axis unit for thousands should be Chinese.");

                // Save the workbook to a temporary file to exercise lifecycle operations
                string tempFile = Path.GetTempFileName();

                // Ensure the path is valid before saving
                if (!string.IsNullOrWhiteSpace(tempFile))
                {
                    workbook.Save(tempFile);
                    // Clean up
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                    }
                }

                Console.WriteLine("All Chinese globalization settings verified successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
                Environment.Exit(1);
            }
        }

        // Simple verification helper that throws on failure
        private static void Verify(bool condition, string message)
        {
            if (!condition)
            {
                throw new InvalidOperationException(message);
            }
        }
    }
}
