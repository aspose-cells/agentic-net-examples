// Title: C# – Batch Export All Workbook Charts with Chinese Localization Using Aspose.Cells
// Description: Loads an Excel workbook, applies a custom ChineseChartGlobalizationSettings to the workbook's globalization settings, iterates every worksheet and chart, assigns a unique name when needed, exports each chart as a PNG image, and optionally saves the localized workbook.
// Keywords: Aspose.Cells | C# chart export | batch export charts | Chinese chart localization | ChartGlobalizationSettings | export Excel charts to PNG | globalization settings Aspose | Excel to image C# | localized chart images
// Common Searches: Aspose.Cells export all charts to PNG | apply Chinese globalization to Excel charts C# | batch process charts in a workbook Aspose | custom ChartGlobalizationSettings example | save each Excel chart as image C#
// Developer Intent: Export every chart in an Excel workbook to PNG files after applying Chinese‑language globalization settings.
// Use Cases: Generate Chinese‑language chart images for reporting dashboards. | Create PNG assets of Excel charts for documentation that requires localized axis and legend labels. | Prepare chart visuals for web or mobile apps with consistent Chinese terminology across all charts.
// AI Prompts: Write C# code that loads an Excel file, sets a custom ChineseChartGlobalizationSettings, iterates all charts, and saves each chart as a PNG image. | Explain how to extend ChineseChartGlobalizationSettings to add translations for additional display units in Aspose.Cells. | Provide a step‑by‑step guide to batch export charts with Chinese localization, then zip the resulting PNG files.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLocalization
{
    // Custom globalization settings that provide Chinese labels for chart elements
    // Loads an Excel workbook, applies a custom ChineseChartGlobalizationSettings to the workbook's globalization settings, iterates every worksheet and chart, assigns a unique name when needed, exports each chart as a PNG image, and optionally saves the localized workbook.
    public class ChineseChartGlobalizationSettings : ChartGlobalizationSettings
    {
        public override string GetAxisUnitName(DisplayUnitType type)
        {
            // Chinese unit names
            return type switch
            {
                DisplayUnitType.Hundreds => "百",
                DisplayUnitType.Thousands => "千",
                DisplayUnitType.TenThousands => "万",
                DisplayUnitType.Millions => "百万",
                DisplayUnitType.Billions => "十亿",
                _ => base.GetAxisUnitName(type)
            };
        }

        public override string GetAxisTitleName()
        {
            return "轴标题";
        }

        public override string GetChartTitleName()
        {
            return "图表标题";
        }

        public override string GetLegendDecreaseName()
        {
            return "递减";
        }

        public override string GetLegendIncreaseName()
        {
            return "递增";
        }

        public override string GetLegendTotalName()
        {
            return "总计";
        }

        public override string GetOtherName()
        {
            return "其他";
        }

        public override string GetSeriesName()
        {
            return "系列";
        }
    }

    class Program
    {
        static void Main()
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Output folder for chart images
            string outputDir = "ChartImages";
            Directory.CreateDirectory(outputDir);

            // Load the workbook (create rule)
            Workbook workbook = new Workbook(inputPath);

            // Apply Chinese globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChineseChartGlobalizationSettings()
            };

            // Iterate through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int chartIdx = 0;
                foreach (Chart chart in sheet.Charts)
                {
                    // Ensure the chart has a name; otherwise generate one
                    string chartName = !string.IsNullOrEmpty(chart.Name)
                        ? chart.Name
                        : $"Sheet{sheet.Index}_Chart{chartIdx}";

                    // Build image file path (PNG format)
                    string imagePath = Path.Combine(outputDir, $"{chartName}.png");

                    // Export chart to image using the rule Chart.ToImage(string, ImageType)
                    chart.ToImage(imagePath, ImageType.Png);

                    chartIdx++;
                }
            }

            // Optionally save the workbook after applying settings (save rule)
            workbook.Save("output_localized.xlsx");
        }
    }
}
