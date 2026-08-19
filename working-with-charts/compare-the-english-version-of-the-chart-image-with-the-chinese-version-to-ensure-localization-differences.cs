// Title: Compare English and Chinese Chart Renderings with Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart, renders it twice—once with default English globalization and once with a custom Chinese ChartGlobalizationSettings—saves each as PNG, then programmatically compares the images to verify that localization changes are applied.
// Keywords: Aspose.Cells | C# | .NET | chart localization | Chinese chart labels | English chart rendering | ChartGlobalizationSettings | render chart to PNG | image comparison C# | globalization settings Aspose | workbook chart rendering | localization testing
// Common Searches: Aspose.Cells render chart in Chinese | C# compare two chart PNG files | How to use ChartGlobalizationSettings with Aspose.Cells | Localization of chart titles Aspose.Cells .NET | Render chart as image with English and Chinese labels
// Developer Intent: Generate English and Chinese versions of a chart image and automatically detect whether the localization settings alter the rendered output.
// Use Cases: Validate that custom Chinese labels appear correctly on axis titles, legends, and titles after applying ChartGlobalizationSettings. | Automate regression tests for chart localization across multiple languages in a CI/CD pipeline. | Create side‑by‑side PNG assets for documentation or UI demos that showcase English vs. Chinese chart presentations.
// AI Prompts: Write a C# function that renders a workbook chart in both English and Chinese using Aspose.Cells and returns true if the PNG files differ. | Generate code to extract the localized chart title, axis titles, and legend entries after applying a custom ChartGlobalizationSettings class. | Create an xUnit test that verifies the Chinese ChartGlobalizationSettings produce the expected label text in the rendered PNG image.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartLocalizationComparison
{
    // Custom globalization settings that provide Chinese labels for chart elements
    // Creates a workbook, adds a column chart, renders it twice—once with default English globalization and once with a custom Chinese ChartGlobalizationSettings—saves each as PNG, then programmatically compares the images to verify that localization changes are applied.
    public class ChineseChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetChartTitleName() => "图表标题";
        public override string GetAxisTitleName() => "轴标题";
        public override string GetOtherName() => "其他";
        public override string GetLegendIncreaseName() => "增加";
        public override string GetLegendDecreaseName() => "减少";
        public override string GetLegendTotalName() => "总计";
        public override string GetSeriesName() => "系列";
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            // Simple mapping for demonstration
            return type switch
            {
                DisplayUnitType.Hundreds => "百",
                DisplayUnitType.Thousands => "千",
                DisplayUnitType.TenThousands => "万",
                _ => base.GetAxisUnitName(type),
            };
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Create workbook and chart --------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set a generic chart title (will be replaced by localization settings)
                chart.Title.Text = "Sample Chart";

                // -------------------- Render English version --------------------
                // Ensure default (English) globalization settings
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings(); // defaults to English

                ImageOrPrintOptions engOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG; no need to set ImageFormat explicitly
                    DefaultEditLanguage = DefaultEditLanguage.English
                };

                SheetRender engRenderer = new SheetRender(sheet, engOptions);
                string engImagePath = "chart_en.png";

                // Render and save English image
                engRenderer.ToImage(0, engImagePath);

                // -------------------- Render Chinese version --------------------
                // Apply Chinese globalization settings
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new ChineseChartGlobalizationSettings()
                };

                ImageOrPrintOptions chiOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG
                    DefaultEditLanguage = DefaultEditLanguage.CJK
                };

                SheetRender chiRenderer = new SheetRender(sheet, chiOptions);
                string chiImagePath = "chart_zh.png";

                // Render and save Chinese image
                chiRenderer.ToImage(0, chiImagePath);

                // -------------------- Compare the two images byte by byte --------------------
                bool imagesAreIdentical = CompareFiles(engImagePath, chiImagePath);
                Console.WriteLine(imagesAreIdentical
                    ? "The English and Chinese chart images are identical."
                    : "The English and Chinese chart images differ (localization applied).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to compare two files byte by byte
        private static bool CompareFiles(string path1, string path2)
        {
            if (!File.Exists(path1) || !File.Exists(path2))
                return false;

            byte[] bytes1 = File.ReadAllBytes(path1);
            byte[] bytes2 = File.ReadAllBytes(path2);

            if (bytes1.Length != bytes2.Length)
                return false;

            for (int i = 0; i < bytes1.Length; i++)
            {
                if (bytes1[i] != bytes2[i])
                    return false;
            }
            return true;
        }
    }
}
