// Title: C# Unit Test for Aspose.Cells ChartChineseSettings – Verify Chinese Title & Legend Strings
// Description: Example code that defines a ChartChineseSettings class inheriting from ChartGlobalizationSettings to supply Chinese text for chart titles and legend actions. Includes self‑contained unit tests that assert each overridden method returns the expected Chinese string and a test that applies the settings to a Workbook, creates a column chart, saves it to a MemoryStream, and verifies the stream is not empty.
// Keywords: Aspose.Cells | .NET | C# | ChartGlobalizationSettings | ChartChineseSettings | unit test | MSTest | NUnit | xUnit | Chinese localization | chart title | legend increase | legend decrease | Workbook save | MemoryStream
// Common Searches: Aspose.Cells unit test chart globalization | How to test ChartChineseSettings in C# | Verify Chinese chart title Aspose.Cells | Custom chart localization .NET | Save workbook with Chinese chart settings
// Developer Intent: Create automated tests that confirm the custom ChartChineseSettings class returns the correct Chinese strings for title and legend, and that a workbook using these settings can be saved without errors.
// Use Cases: Validate that GetChartTitleName returns "图表标题" for proper title localization. | Ensure GetLegendIncreaseName and GetLegendDecreaseName return "增加" and "减少" respectively, guaranteeing correct legend labeling. | Confirm a workbook configured with ChartChineseSettings can add a chart and be saved to a non‑empty MemoryStream.
// AI Prompts: Generate MSTest methods that assert ChartChineseSettings overrides for title, legend increase, and legend decrease strings. | Write an xUnit test that creates a Workbook, sets GlobalizationSettings.ChartSettings to ChartChineseSettings, adds a column chart, saves to a MemoryStream, and checks the stream length is greater than zero. | Provide a reusable test helper that compares expected Chinese strings with actual values from ChartChineseSettings and throws a descriptive exception on mismatch.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Custom globalization settings that provide Chinese strings
    // Example code that defines a ChartChineseSettings class inheriting from ChartGlobalizationSettings to supply Chinese text for chart titles and legend actions. Includes self‑contained unit tests that assert each overridden method returns the expected Chinese string and a test that applies the settings to a Workbook, creates a column chart, saves it to a MemoryStream, and verifies the stream is not empty.
    public class ChartChineseSettings : ChartGlobalizationSettings
    {
        public override string GetChartTitleName() => "图表标题";
        public override string GetLegendIncreaseName() => "增加";
        public override string GetLegendDecreaseName() => "减少";
    }

    public static class Tests
    {
        public static void RunAll()
        {
            GetChartTitleName_ReturnsExpectedChineseString();
            GetLegendIncreaseName_ReturnsExpectedChineseString();
            GetLegendDecreaseName_ReturnsExpectedChineseString();
            Workbook_WithChineseChartSettings_SavesSuccessfully();
            Console.WriteLine("All tests passed.");
        }

        private static void AssertEqual(string expected, string actual, string testName)
        {
            if (!expected.Equals(actual))
                throw new Exception($"Assertion failed in {testName}: expected '{expected}', got '{actual}'.");
        }

        private static void GetChartTitleName_ReturnsExpectedChineseString()
        {
            var settings = new ChartChineseSettings();
            AssertEqual("图表标题", settings.GetChartTitleName(), nameof(GetChartTitleName_ReturnsExpectedChineseString));
        }

        private static void GetLegendIncreaseName_ReturnsExpectedChineseString()
        {
            var settings = new ChartChineseSettings();
            AssertEqual("增加", settings.GetLegendIncreaseName(), nameof(GetLegendIncreaseName_ReturnsExpectedChineseString));
        }

        private static void GetLegendDecreaseName_ReturnsExpectedChineseString()
        {
            var settings = new ChartChineseSettings();
            AssertEqual("减少", settings.GetLegendDecreaseName(), nameof(GetLegendDecreaseName_ReturnsExpectedChineseString));
        }

        private static void Workbook_WithChineseChartSettings_SavesSuccessfully()
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Apply custom Chinese globalization settings
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new ChartChineseSettings()
                };

                // Add a simple chart
                Worksheet sheet = workbook.Worksheets[0];
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("{10,20,30}", true);
                chart.Title.Text = "Demo";

                // Save to memory stream
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    if (ms.Length == 0)
                        throw new Exception("Memory stream is empty after saving workbook.");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Workbook_WithChineseChartSettings_SavesSuccessfully failed.", ex);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                Tests.RunAll();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
