// Title: Compare English and Chinese Chart Localization with Aspose.Cells for .NET
// Description: This C# sample creates a workbook, adds a column chart, renders an English PNG using the default ChartGlobalizationSettings, then applies a custom ChineseChartGlobalizationSettings to localize titles, axis and legend labels, renders a Chinese PNG, prints both sets of strings to the console, and saves the workbook with the localized chart.
// Keywords: Aspose.Cells chart localization | C# chart globalization | English Chinese chart comparison | ChartGlobalizationSettings example | render chart PNG Aspose.Cells
// Common Searches: Aspose.Cells how to localize chart labels | render chart image in Chinese with Aspose.Cells | compare English and Chinese chart screenshots .NET | custom ChartGlobalizationSettings C# | Aspose.Cells chart localization tutorial
// Developer Intent: Generate side‑by‑side English and Chinese chart images and verify that all chart elements are correctly localized using Aspose.Cells.
// Use Cases: Create a column chart from sample data and export an English PNG with default globalization. | Apply a custom ChineseChartGlobalizationSettings to replace chart titles, axis titles, and legend entries with Chinese text. | Export a Chinese PNG, output both language strings for validation, and save the workbook containing the localized chart.
// AI Prompts: Write a C# example that defines a FrenchChartGlobalizationSettings class and renders English and French chart images with Aspose.Cells. | Provide code to programmatically compare two chart PNG files (English vs. Chinese) pixel by pixel and generate a diff report.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartLocalizationComparison
{
    // Custom globalization settings that return Chinese strings for chart elements
    // This C# sample creates a workbook, adds a column chart, renders an English PNG using the default ChartGlobalizationSettings, then applies a custom ChineseChartGlobalizationSettings to localize titles, axis and legend labels, renders a Chinese PNG, prints both sets of strings to the console, and saves the workbook with the localized chart.
    public class ChineseChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetChartTitleName() => "图表标题"; // "Chart Title" in Chinese
        public override string GetAxisTitleName() => "轴标题"; // "Axis Title" in Chinese
        public override string GetLegendIncreaseName() => "增加"; // "Increase" in Chinese
        public override string GetLegendDecreaseName() => "减少"; // "Decrease" in Chinese
        public override string GetLegendTotalName() => "总计"; // "Total" in Chinese
        public override string GetOtherName() => "其他"; // "Other" in Chinese
        public override string GetSeriesName() => "系列"; // "Series" in Chinese
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Common workbook setup --------------------
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------- English version --------------------
                // Use default (English) globalization settings
                ChartGlobalizationSettings engSettings = new ChartGlobalizationSettings();
                chart.Title.Text = engSettings.GetChartTitleName();

                // Render the English chart image
                ImageOrPrintOptions engOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG; no need to set ImageFormat explicitly
                    DefaultEditLanguage = DefaultEditLanguage.English
                };
                SheetRender engRenderer = new SheetRender(sheet, engOptions);
                engRenderer.ToImage(0, "Chart_English.png");

                // Output English localization strings for verification
                Console.WriteLine("English Chart Title: " + engSettings.GetChartTitleName());
                Console.WriteLine("English Axis Title: " + engSettings.GetAxisTitleName());
                Console.WriteLine("English Legend Increase: " + engSettings.GetLegendIncreaseName());
                Console.WriteLine("English Legend Decrease: " + engSettings.GetLegendDecreaseName());
                Console.WriteLine("English Legend Total: " + engSettings.GetLegendTotalName());
                Console.WriteLine("English Other Label: " + engSettings.GetOtherName());

                // -------------------- Chinese version --------------------
                // Apply custom Chinese globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new ChineseChartGlobalizationSettings()
                };

                // Retrieve the Chinese settings (they are now attached to the workbook)
                ChartGlobalizationSettings chiSettings = workbook.Settings.GlobalizationSettings.ChartSettings;

                // Update chart title with Chinese name
                chart.Title.Text = chiSettings.GetChartTitleName();

                // Render the Chinese chart image
                ImageOrPrintOptions chiOptions = new ImageOrPrintOptions
                {
                    // Default image format is PNG; no need to set ImageFormat explicitly
                    DefaultEditLanguage = DefaultEditLanguage.CJK
                };
                SheetRender chiRenderer = new SheetRender(sheet, chiOptions);
                chiRenderer.ToImage(0, "Chart_Chinese.png");

                // Output Chinese localization strings for verification
                Console.WriteLine("Chinese Chart Title: " + chiSettings.GetChartTitleName());
                Console.WriteLine("Chinese Axis Title: " + chiSettings.GetAxisTitleName());
                Console.WriteLine("Chinese Legend Increase: " + chiSettings.GetLegendIncreaseName());
                Console.WriteLine("Chinese Legend Decrease: " + chiSettings.GetLegendDecreaseName());
                Console.WriteLine("Chinese Legend Total: " + chiSettings.GetLegendTotalName());
                Console.WriteLine("Chinese Other Label: " + chiSettings.GetOtherName());

                // -------------------- Save workbook --------------------
                // Save the workbook (contains the chart with Chinese settings applied)
                string outputPath = "ChartLocalizationComparison.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
