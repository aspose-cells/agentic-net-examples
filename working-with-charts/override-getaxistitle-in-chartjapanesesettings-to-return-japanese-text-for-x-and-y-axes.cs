// Title: Override GetAxisTitleName to Set Japanese X‑ and Y‑Axis Titles in Aspose.Cells Charts (C#)
// Description: Demonstrates how to create a custom ChartJapaneseSettings class that inherits SettableChartGlobalizationSettings, overrides GetAxisTitleName to return Japanese text, and applies the titles to the CategoryAxis and ValueAxis of a column chart before saving the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# chart localization | Japanese axis titles | SettableChartGlobalizationSettings | GetAxisTitleName override | Excel chart globalization | column chart Japanese labels | regional Excel reporting
// Common Searches: Aspose.Cells set Japanese axis title | override GetAxisTitleName C# | chart globalization settings Aspose.Cells | how to localize chart axes in Excel using Aspose | C# example Japanese chart titles Aspose.Cells
// Developer Intent: Implement a custom globalization class to display Japanese titles on both axes of an Aspose.Cells chart.
// Use Cases: Generate sales dashboards with axis labels in Japanese for local stakeholders. | Standardize chart axis localization across multiple workbooks in a multinational application. | Extend the overridden method to provide distinct Japanese strings for category and value axes.
// AI Prompts: Write C# code that creates a SettableChartGlobalizationSettings subclass overriding GetAxisTitleName to return separate Japanese strings for the X‑axis and Y‑axis. | Explain step‑by‑step how to apply a custom globalization settings class to an Aspose.Cells chart and enable axis title visibility. | Provide a complete example that sets Japanese titles on a column chart’s CategoryAxis and ValueAxis, then saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Custom globalization settings that provide Japanese axis titles
// Demonstrates how to create a custom ChartJapaneseSettings class that inherits SettableChartGlobalizationSettings, overrides GetAxisTitleName to return Japanese text, and applies the titles to the CategoryAxis and ValueAxis of a column chart before saving the workbook as an XLSX file.
class ChartJapaneseSettings : SettableChartGlobalizationSettings
{
    // Override to return Japanese text for axis titles
    public override string GetAxisTitleName()
    {
        // You can customize this string as needed, e.g., "X軸" or "Y軸"
        return "軸タイトル";
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("カテゴリ");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["B1"].PutValue("売上");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply custom Japanese globalization settings
        ChartJapaneseSettings japaneseSettings = new ChartJapaneseSettings();

        // Set axis titles using the overridden method
        chart.CategoryAxis.Title.Text = japaneseSettings.GetAxisTitleName(); // X‑axis
        chart.ValueAxis.Title.Text = japaneseSettings.GetAxisTitleName();    // Y‑axis

        // Make titles visible
        chart.CategoryAxis.Title.IsVisible = true;
        chart.ValueAxis.Title.IsVisible = true;

        // Save the workbook
        workbook.Save("ChartWithJapaneseAxisTitles.xlsx");
    }
}
