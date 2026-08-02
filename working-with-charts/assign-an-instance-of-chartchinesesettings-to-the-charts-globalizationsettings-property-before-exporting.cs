// Title: Set Chinese Globalization Settings for Aspose.Cells Charts in C# Before Export
// Description: Demonstrates how to create a custom ChartChineseSettings class derived from SettableChartGlobalizationSettings, assign Chinese names to chart titles, axis titles, units, legends and series, apply the settings to a workbook's GlobalizationSettings.ChartSettings, and export the workbook as an Excel file.
// Keywords: Aspose.Cells C# chart globalization | Chinese chart localization | SettableChartGlobalizationSettings example | Excel chart export Chinese labels | Aspose.Cells chart titles Chinese | axis unit Chinese Aspose.Cells | chart legend Chinese Aspose | GitHub Aspose.Cells chart localization | globalization settings for Excel charts
// Common Searches: how to localize Aspose.Cells chart labels to Chinese | apply custom ChartGlobalizationSettings in C# | Aspose.Cells chart Chinese axis units example | set Chinese legend names for Excel charts using Aspose | C# code for chart globalization settings Aspose.Cells
// Developer Intent: Apply a custom ChartChineseSettings object to a workbook’s GlobalizationSettings.ChartSettings so that all chart elements are rendered with Chinese terminology before saving.
// Use Cases: Generate Excel reports with column charts that display titles, axis labels, and legends in Chinese for regional users. | Standardize chart unit displays (hundreds, thousands, ten‑thousands) using Chinese characters across multiple workbooks. | Create a reusable localization class for Aspose.Cells charts that can be applied to any workbook prior to export.
// AI Prompts: Show how to assign ChartChineseSettings to a single chart instead of the whole workbook in Aspose.Cells C#. | Generate C# code that configures Japanese chart globalization settings using SettableChartGlobalizationSettings. | Explain how to read, modify, and re‑apply ChartGlobalizationSettings after loading an existing Excel file with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartChineseSettingsDemo
{
    // Custom globalization settings for charts with Chinese labels
    // Demonstrates how to create a custom ChartChineseSettings class derived from SettableChartGlobalizationSettings, assign Chinese names to chart titles, axis titles, units, legends and series, apply the settings to a workbook's GlobalizationSettings.ChartSettings, and export the workbook as an Excel file.
    public class ChartChineseSettings : SettableChartGlobalizationSettings
    {
        public ChartChineseSettings()
        {
            // Set Chinese names for various chart elements
            SetChartTitleName("图表标题");
            SetAxisTitleName("轴标题");
            SetAxisUnitName(DisplayUnitType.Hundreds, "百");
            SetAxisUnitName(DisplayUnitType.Thousands, "千");
            SetAxisUnitName(DisplayUnitType.TenThousands, "万");
            SetLegendIncreaseName("增加");
            SetLegendDecreaseName("减少");
            SetLegendTotalName("合计");
            SetOtherName("其他");
            SetSeriesName("系列");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("类别");
            sheet.Cells["A2"].PutValue("第一季度");
            sheet.Cells["A3"].PutValue("第二季度");
            sheet.Cells["A4"].PutValue("第三季度");
            sheet.Cells["B1"].PutValue("数值");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply custom Chinese globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChartChineseSettings()
            };

            // Save the workbook (export)
            workbook.Save("ChartWithChineseSettings.xlsx");
        }
    }
}
