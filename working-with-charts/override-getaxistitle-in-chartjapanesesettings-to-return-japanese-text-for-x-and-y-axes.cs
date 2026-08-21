// Title: C# Aspose.Cells – Override GetAxisTitleName to set Japanese X‑axis and Y‑axis titles
// Description: This sample creates a custom ChartJapaneseSettings class that inherits SettableChartGlobalizationSettings and overrides GetAxisTitleName to return a Japanese label. The returned text is assigned to both the CategoryAxis and ValueAxis of a column chart, the titles are made visible, and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# chart localization | Japanese axis titles | SettableChartGlobalizationSettings | GetAxisTitleName override | column chart axis label | X軸 Y軸 | globalization settings | chart title visibility | GitHub example
// Common Searches: override GetAxisTitleName Aspose.Cells C# | Japanese axis labels for Aspose.Cells chart | custom chart globalization settings Aspose | set axis title text programmatically Aspose.Cells | localize chart axes to Japanese in .NET
// Developer Intent: Provide Japanese text for the X and Y axes of an Aspose.Cells chart by customizing the globalization settings class.
// Use Cases: Create sales dashboards for Japanese markets with correctly labeled chart axes. | Reuse a single globalization class to apply consistent Japanese axis titles across multiple charts in a workbook. | Programmatically toggle axis title visibility to improve readability in localized reports.
// AI Prompts: Write C# code that overrides SettableChartGlobalizationSettings.GetAxisTitleName to return "X軸" for the category axis and "Y軸" for the value axis, then apply them to a chart. | Explain how to modify ChartJapaneseSettings so each axis receives a distinct Japanese label while keeping a single overridden method. | Generate a step‑by‑step tutorial for localizing all chart axis titles to Japanese in an Aspose.Cells workbook using a custom globalization class.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Custom globalization settings that provide Japanese axis titles
    // This sample creates a custom ChartJapaneseSettings class that inherits SettableChartGlobalizationSettings and overrides GetAxisTitleName to return a Japanese label. The returned text is assigned to both the CategoryAxis and ValueAxis of a column chart, the titles are made visible, and the workbook is saved as an XLSX file.
    public class ChartJapaneseSettings : SettableChartGlobalizationSettings
    {
        // Override the method to return Japanese text for axis titles
        public override string GetAxisTitleName()
        {
            // "X軸" and "Y軸" are commonly used Japanese labels.
            // Since the method does not differentiate between X and Y,
            // we return a generic Japanese term that can be used for both.
            return "軸";
        }
    }

    public class ChartJapaneseSettingsDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("カテゴリ");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("値");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Use the custom Japanese settings to set axis titles
            ChartJapaneseSettings japaneseSettings = new ChartJapaneseSettings();
            string axisTitle = japaneseSettings.GetAxisTitleName();

            // Apply the Japanese axis title to both X (Category) and Y (Value) axes
            chart.CategoryAxis.Title.Text = axisTitle; // X‑axis
            chart.ValueAxis.Title.Text = axisTitle;    // Y‑axis

            // Optionally make the titles visible
            chart.CategoryAxis.Title.IsVisible = true;
            chart.ValueAxis.Title.IsVisible = true;

            // Save the workbook
            workbook.Save("ChartJapaneseSettingsDemo.xlsx");
        }
    }

    // Entry point
    class Program
    {
        static void Main()
        {
            ChartJapaneseSettingsDemo.Run();
        }
    }
}
