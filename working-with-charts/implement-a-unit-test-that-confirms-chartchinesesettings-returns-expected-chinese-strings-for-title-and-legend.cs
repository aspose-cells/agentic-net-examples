using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTests
{
    // Custom globalization settings that provide Chinese strings
    public class ChineseChartSettings : ChartGlobalizationSettings
    {
        public override string GetChartTitleName() => "图表标题";

        public override string GetLegendIncreaseName() => "增加";

        public override string GetLegendDecreaseName() => "减少";

        public override string GetLegendTotalName() => "总计";

        // Other members use base implementation
    }

    public class ChartChineseSettingsTests
    {
        public void Run()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                Workbook workbook = new Workbook();

                // Apply the custom Chinese globalization settings
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new ChineseChartSettings()
                };

                // Retrieve the applied settings
                ChartGlobalizationSettings settings = workbook.Settings.GlobalizationSettings.ChartSettings;

                // Verify that the Chinese strings are returned as expected
                if (settings.GetChartTitleName() != "图表标题" ||
                    settings.GetLegendIncreaseName() != "增加" ||
                    settings.GetLegendDecreaseName() != "减少" ||
                    settings.GetLegendTotalName() != "总计")
                {
                    throw new Exception("Chinese globalization settings validation failed.");
                }

                // Save the workbook to a memory stream (lifecycle save) to ensure no file I/O errors
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Save(ms, SaveFormat.Xlsx);
                    // Verify that the stream contains data
                    if (ms.Length == 0)
                        throw new Exception("Workbook was saved but the stream is empty.");
                }

                Console.WriteLine("Test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during test execution: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Example of loading a template workbook if it exists
            string templatePath = "template.xlsx";
            Workbook workbook;

            if (File.Exists(templatePath))
            {
                try
                {
                    workbook = new Workbook(templatePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load template '{templatePath}': {ex.Message}");
                    workbook = new Workbook(); // fallback to a new workbook
                }
            }
            else
            {
                workbook = new Workbook(); // no template, create a new workbook
            }

            // Run the Chinese globalization settings test
            var test = new ChartChineseSettingsTests();
            test.Run();
        }
    }
}