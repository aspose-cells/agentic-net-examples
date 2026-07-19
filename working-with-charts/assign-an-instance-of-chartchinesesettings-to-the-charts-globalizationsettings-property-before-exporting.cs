// Title: C# – Apply Chinese Globalization Settings to Aspose.Cells Charts Before Export
// Description: Shows how to subclass SettableChartGlobalizationSettings to define Chinese labels for chart titles, series, legends and axis units, attach the subclass to workbook.Settings.GlobalizationSettings.ChartSettings, create a column chart with sample data, and save the workbook as an XLSX file.
// Keywords: Aspose.Cells C# chart globalization | SettableChartGlobalizationSettings Chinese | custom chart labels Aspose.Cells | chart axis unit Chinese Aspose | export workbook with localized chart | globalizationsettings chart Aspose.Cells | Chinese localization for Excel charts | Aspose.Cells chart title translation | C# Excel chart localization
// Common Searches: how to set Chinese labels for Aspose.Cells charts | assign custom ChartGlobalizationSettings in .NET | Aspose.Cells chart axis unit names Chinese | localize Aspose.Cells chart elements | globalize chart titles and legends in Excel using Aspose
// Developer Intent: Attach a ChartChineseSettings object to the workbook’s GlobalizationSettings.ChartSettings so that exported charts display Chinese titles, series names, legends, and axis units.
// Use Cases: Generate business reports with all chart text in Chinese for Chinese‑language audiences. | Display financial or statistical axis units using Chinese characters (百, 千, 万). | Apply a single Chinese globalization configuration to every chart in a workbook to ensure consistent localization. | Create multilingual Excel dashboards by swapping different ChartGlobalizationSettings subclasses.
// AI Prompts: Write C# code that creates a SettableChartGlobalizationSettings subclass for Japanese chart localization and applies it to an Aspose.Cells workbook. | Explain how to set different ChartGlobalizationSettings for individual charts instead of using the workbook‑level setting. | Show how to load an existing XLSX file, read its current ChartGlobalizationSettings, modify a label, and save the changes.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartGlobalizationDemo
{
    // Custom globalization settings for charts with Chinese labels
    // Shows how to subclass SettableChartGlobalizationSettings to define Chinese labels for chart titles, series, legends and axis units, attach the subclass to workbook.Settings.GlobalizationSettings.ChartSettings, create a column chart with sample data, and save the workbook as an XLSX file.
    public class ChartChineseSettings : SettableChartGlobalizationSettings
    {
        public ChartChineseSettings()
        {
            // Set Chinese names for common chart elements
            SetChartTitleName("图表标题");
            SetSeriesName("系列");
            SetLegendIncreaseName("增加");
            SetLegendDecreaseName("减少");
            SetLegendTotalName("合计");
            SetOtherName("其他");

            // Set Chinese unit names for axis display units
            SetAxisUnitName(DisplayUnitType.Hundreds, "百");
            SetAxisUnitName(DisplayUnitType.Thousands, "千");
            SetAxisUnitName(DisplayUnitType.TenThousands, "万");
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("类别");
            sheet.Cells["A2"].PutValue("第一季度");
            sheet.Cells["A3"].PutValue("第二季度");
            sheet.Cells["A4"].PutValue("第三季度");
            sheet.Cells["B1"].PutValue("数值");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(2500);
            sheet.Cells["B4"].PutValue(3700);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "销售趋势";

            // Apply custom Chinese globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = new ChartChineseSettings()
            };

            // Export the workbook (e.g., to XLSX)
            workbook.Save("ChartWithChineseGlobalization.xlsx");
        }
    }
}
