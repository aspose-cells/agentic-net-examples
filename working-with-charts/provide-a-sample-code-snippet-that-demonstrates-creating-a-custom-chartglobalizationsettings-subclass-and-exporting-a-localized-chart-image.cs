// Title: Aspose.Cells .NET – Subclass ChartGlobalizationSettings to localize chart labels and export PNG
// Description: Demonstrates how to create a MyChartGlobalizationSettings class that overrides axis unit names, chart title, and legend texts with Chinese strings, apply it to a workbook, build a column chart, set a display unit, and export the chart as a PNG image while also saving the Excel file.
// Keywords: Aspose.Cells | C# chart localization | ChartGlobalizationSettings subclass | custom axis unit names | Chinese chart labels | export chart to PNG | Excel chart image generation | globalization settings Aspose | .NET chart rendering
// Common Searches: Aspose.Cells custom ChartGlobalizationSettings example | localize chart axis units in C# Aspose.Cells | export Aspose.Cells chart as PNG with Chinese labels | how to change chart legend text in Aspose.Cells .NET | apply globalization settings to Excel chart programmatically
// Developer Intent: The developer wants to implement a custom globalization class to replace default chart text with Chinese terms and generate a localized chart image.
// Use Cases: Produce Chinese‑language Excel reports with correctly localized chart axes and titles. | Automate creation of web‑ready chart images (PNG) that match regional terminology. | Save both the workbook and a localized chart snapshot for distribution to international stakeholders.
// AI Prompts: Generate C# code that subclasses ChartGlobalizationSettings to use French axis unit names and exports the chart as a JPEG. | Explain how to apply a custom ChartGlobalizationSettings instance to all charts in an Aspose.Cells workbook. | Write a try‑catch block that logs chart image export failures to a log file using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartLocalizationDemo
{
    // Custom globalization settings for charts
    // Demonstrates how to create a MyChartGlobalizationSettings class that overrides axis unit names, chart title, and legend texts with Chinese strings, apply it to a workbook, build a column chart, set a display unit, and export the chart as a PNG image while also saving the Excel file.
    public class MyChartGlobalizationSettings : ChartGlobalizationSettings
    {
        // Override axis unit names to provide localized strings
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            return type switch
            {
                DisplayUnitType.Hundreds => "百",          // Chinese for "hundreds"
                DisplayUnitType.Thousands => "千",          // Chinese for "thousands"
                DisplayUnitType.TenThousands => "万",      // Chinese for "ten‑thousands"
                _ => base.GetAxisUnitName(type),
            };
        }

        // Override chart title name
        public override string GetChartTitleName() => "本地化图表标题";   // "Localized Chart Title"

        // Override legend increase/decrease names
        public override string GetLegendIncreaseName() => "增加";
        public override string GetLegendDecreaseName() => "减少";
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("类别");
                sheet.Cells["A2"].PutValue("第一季");
                sheet.Cells["A3"].PutValue("第二季");
                sheet.Cells["A4"].PutValue("第三季");
                sheet.Cells["B1"].PutValue("销量");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(1500);
                sheet.Cells["B4"].PutValue(1800);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set a display unit to see the custom unit name in action
                chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands;
                chart.ValueAxis.IsDisplayUnitLabelShown = true;

                // Apply the custom globalization settings to the workbook
                wb.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = new MyChartGlobalizationSettings()
                };

                // Apply the custom title
                chart.Title.Text = wb.Settings.GlobalizationSettings.ChartSettings.GetChartTitleName();

                // Configure image export options (default format is PNG)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

                // Export the chart as an image (PNG)
                try
                {
                    chart.ToImage("LocalizedChart.png", imgOptions);
                }
                catch (Exception imgEx)
                {
                    Console.WriteLine($"Image export error: {imgEx.Message}");
                }

                // Save the workbook to see the chart inside Excel
                wb.Save("LocalizedChartWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
