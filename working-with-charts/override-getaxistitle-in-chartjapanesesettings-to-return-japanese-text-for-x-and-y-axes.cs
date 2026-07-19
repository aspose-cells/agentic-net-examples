// Title: Aspose.Cells .NET: Override GetAxisTitleName to set Japanese X‑ and Y‑axis titles in a column chart
// Description: Demonstrates how to create a ChartJapaneseSettings class that inherits SettableChartGlobalizationSettings and overrides GetAxisTitleName to return the Japanese word "軸". The example builds a workbook, adds sample data, inserts a column chart, applies the custom axis titles to both CategoryAxis and ValueAxis, makes them visible, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# chart localization | Japanese axis titles | SettableChartGlobalizationSettings | GetAxisTitleName override | .NET Excel chart customization | column chart Japanese labels | globalization settings Aspose
// Common Searches: Aspose.Cells set Japanese axis title .NET | override GetAxisTitleName for chart localization | how to display Japanese text on chart axes in C# | custom chart globalization settings Aspose.Cells | column chart with Japanese X and Y axis labels
// Developer Intent: Create a subclass of SettableChartGlobalizationSettings that returns Japanese text for axis titles and apply it to a chart's CategoryAxis and ValueAxis.
// Use Cases: Define ChartJapaneseSettings that overrides GetAxisTitleName to supply a Japanese string. | Assign the overridden title to chart.CategoryAxis.Title.Text and chart.ValueAxis.Title.Text. | Enable axis title visibility and generate an Excel file with localized axis labels. | Reuse the custom globalization class for other chart types requiring Japanese titles.
// AI Prompts: Generate C# code that overrides GetAxisTitleName to return different Japanese strings for X‑axis and Y‑axis in Aspose.Cells. | Show how to apply a custom ChartJapaneseSettings class to a column chart and ensure the axis titles appear in Japanese. | Explain step‑by‑step how to localize chart axis titles for multiple languages using SettableChartGlobalizationSettings in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Custom globalization settings for Japanese axis titles
    // Demonstrates how to create a ChartJapaneseSettings class that inherits SettableChartGlobalizationSettings and overrides GetAxisTitleName to return the Japanese word "軸". The example builds a workbook, adds sample data, inserts a column chart, applies the custom axis titles to both CategoryAxis and ValueAxis, makes them visible, and saves the file as an Excel workbook.
    public class ChartJapaneseSettings : SettableChartGlobalizationSettings
    {
        // Override to provide Japanese axis title text
        public override string GetAxisTitleName()
        {
            // Returns a generic Japanese word for "Axis"
            // You can customize this further to differentiate X/Y if needed
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

            // Populate sample data
            sheet.Cells["A1"].PutValue("カテゴリ");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["B1"].PutValue("売上");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["B3"].PutValue(200);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Use the custom Japanese settings for axis titles
            ChartJapaneseSettings japaneseSettings = new ChartJapaneseSettings();

            // Apply Japanese axis titles
            chart.CategoryAxis.Title.Text = japaneseSettings.GetAxisTitleName(); // X‑axis
            chart.ValueAxis.Title.Text = japaneseSettings.GetAxisTitleName();    // Y‑axis

            // Make titles visible
            chart.CategoryAxis.Title.IsVisible = true;
            chart.ValueAxis.Title.IsVisible = true;

            // Save the workbook
            workbook.Save("ChartJapaneseSettingsDemo.xlsx");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            ChartJapaneseSettingsDemo.Run();
        }
    }
}
