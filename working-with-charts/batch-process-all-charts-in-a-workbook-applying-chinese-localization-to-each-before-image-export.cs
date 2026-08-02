// Title: Batch Export Workbook Charts with Chinese Localization via Aspose.Cells for .NET
// Description: Loads an Excel workbook, configures Chinese chart globalization (titles, series, legend, axis units), applies the settings to the entire workbook, iterates every worksheet to export each chart as a uniquely‑named PNG image, and optionally saves the localized workbook.
// Keywords: Aspose.Cells | .NET | C# | chart export | PNG image | Chinese localization | chart globalization settings | batch chart processing | Excel chart image | SetChartTitleName | DisplayUnitType
// Common Searches: Aspose.Cells export all charts to PNG | How to apply Chinese labels to Excel charts with Aspose.Cells | Batch process charts in a workbook using .NET | Set Chinese axis unit names in Aspose.Cells charts | Localize Excel chart titles to 中文 with Aspose.Cells
// Developer Intent: Apply Chinese chart globalization to every chart in a workbook and generate PNG images for each chart.
// Use Cases: Create a set of Chinese‑language chart images for a reporting dashboard. | Produce PNG assets of all workbook charts for a presentation that requires localized labels. | Save a workbook with Chinese chart captions while also providing image copies for downstream systems.
// AI Prompts: Show how to export the charts as JPEG with a custom quality level. | Add a semi‑transparent watermark to each exported chart image using Aspose.Cells. | Log the total number of charts processed after applying the Chinese globalization settings.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartLocalization
{
    // Loads an Excel workbook, configures Chinese chart globalization (titles, series, legend, axis units), applies the settings to the entire workbook, iterates every worksheet to export each chart as a uniquely‑named PNG image, and optionally saves the localized workbook.
    class Program
    {
        static void Main()
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Output folder for chart images
            string outputDir = "ChartImages";
            Directory.CreateDirectory(outputDir);

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Create customizable Chinese globalization settings
            SettableChartGlobalizationSettings chineseSettings = new SettableChartGlobalizationSettings();
            chineseSettings.SetChartTitleName("图表标题");
            chineseSettings.SetSeriesName("系列");
            chineseSettings.SetLegendIncreaseName("增加");
            chineseSettings.SetLegendDecreaseName("减少");
            chineseSettings.SetOtherName("其他");
            chineseSettings.SetAxisUnitName(DisplayUnitType.Hundreds, "百");
            chineseSettings.SetAxisUnitName(DisplayUnitType.Thousands, "千");
            chineseSettings.SetAxisUnitName(DisplayUnitType.TenThousands, "万");

            // Apply the settings to the workbook's globalization settings
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = chineseSettings
            };

            // Iterate through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int chartIdx = 0;
                foreach (Chart chart in sheet.Charts)
                {
                    // Build a unique file name for each chart image
                    string imageFile = Path.Combine(outputDir,
                        $"Sheet{sheet.Index}_Chart{chartIdx}.png");

                    // Export chart to PNG image
                    chart.ToImage(imageFile, ImageType.Png);

                    chartIdx++;
                }
            }

            // Optionally save the workbook with applied settings
            workbook.Save("LocalizedWorkbook.xlsx");
        }
    }
}
