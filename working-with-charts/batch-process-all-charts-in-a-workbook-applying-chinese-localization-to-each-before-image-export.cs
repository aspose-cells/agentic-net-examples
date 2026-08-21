// Title: Batch apply Chinese chart globalization and export all workbook charts as PNG using Aspose.Cells for .NET
// Description: Loads an Excel workbook, creates Chinese chart globalization settings via SettableChartGlobalizationSettings, assigns them to the workbook, recalculates each chart, and exports every chart from all worksheets to uniquely named PNG files; optionally saves the workbook with the localized settings.
// Keywords: Aspose.Cells | .NET | chart globalization | Chinese localization | batch chart export | export chart to PNG | SettableChartGlobalizationSettings | Excel chart image | globalization settings | chart to image API
// Common Searches: Aspose.Cells Chinese chart globalization | export all charts to PNG Aspose.Cells | batch export Excel charts .NET | set chart titles in Chinese Aspose.Cells | globalize chart labels for Chinese workbook | save Excel charts as images with Aspose.Cells
// Developer Intent: Apply Chinese globalization to every chart in a workbook and generate PNG images for each chart.
// Use Cases: Produce localized chart images for a Chinese-language reporting portal by batch exporting charts after applying Chinese labels. | Create PNG assets of all workbook charts for documentation while preserving Chinese terminology in titles and axes. | Save a copy of the workbook with Chinese chart labels and also generate separate chart images for use in mobile or web applications.
// AI Prompts: Show how to change the export format from PNG to JPEG while keeping the localization settings. | Explain how to assign different globalization settings per worksheet or per chart type. | Provide robust error handling for missing output directories, unsupported chart types, and permission issues during batch export.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, creates Chinese chart globalization settings via SettableChartGlobalizationSettings, assigns them to the workbook, recalculates each chart, and exports every chart from all worksheets to uniquely named PNG files; optionally saves the workbook with the localized settings.
class BatchChartLocalizationAndExport
{
    static void Main()
    {
        // Input workbook path
        string workbookPath = "input.xlsx";

        // Output folder for chart images
        string outputFolder = "ChartImages";
        Directory.CreateDirectory(outputFolder);

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Create customizable Chinese globalization settings
        var chineseSettings = new SettableChartGlobalizationSettings();
        chineseSettings.SetChartTitleName("图表标题");
        chineseSettings.SetSeriesName("系列");
        chineseSettings.SetLegendIncreaseName("增加");
        chineseSettings.SetLegendDecreaseName("减少");
        chineseSettings.SetOtherName("其他");
        chineseSettings.SetAxisTitleName("轴标题");
        chineseSettings.SetAxisUnitName(DisplayUnitType.Hundreds, "百");
        chineseSettings.SetAxisUnitName(DisplayUnitType.Thousands, "千");
        chineseSettings.SetAxisUnitName(DisplayUnitType.TenThousands, "万");

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = chineseSettings
        };

        // Iterate through all worksheets and their charts
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            for (int i = 0; i < sheet.Charts.Count; i++)
            {
                Chart chart = sheet.Charts[i];

                // Ensure chart layout is up‑to‑date
                chart.Calculate();

                // Build a unique file name for each chart image
                string imageFile = Path.Combine(
                    outputFolder,
                    $"{sheet.Name}_Chart{i + 1}.png");

                // Export chart as PNG image
                chart.ToImage(imageFile, ImageType.Png);
                Console.WriteLine($"Exported chart to: {imageFile}");
            }
        }

        // Optionally save the workbook with applied settings
        workbook.Save("output_with_chinese_localization.xlsx");
    }
}
