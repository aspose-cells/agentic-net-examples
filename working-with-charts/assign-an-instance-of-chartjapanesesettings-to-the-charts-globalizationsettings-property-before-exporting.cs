// Title: Aspose.Cells .NET – Apply Japanese Axis Units with ChartJapaneseSettings (GlobalizationSettings)
// Description: This example shows how to subclass ChartGlobalizationSettings, override GetAxisUnitName to return Japanese characters for hundreds, thousands and ten‑thousands, and assign the custom ChartJapaneseSettings to workbook.Settings.GlobalizationSettings.ChartSettings before saving the workbook as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | ChartGlobalizationSettings | ChartJapaneseSettings | Japanese axis labels | Excel chart localization | GlobalizationSettings | DisplayUnitType | Excel export
// Common Searches: Aspose.Cells set Japanese chart axis labels | How to customize chart units in Aspose.Cells .NET | ChartGlobalizationSettings example C# | Assign ChartJapaneseSettings to workbook | Localize Excel chart axis with Aspose.Cells
// Developer Intent: The developer wants to localize chart axis unit names to Japanese by attaching a custom ChartJapaneseSettings object to the workbook’s globalization settings before exporting.
// Use Cases: Create financial charts for Japanese reports where axis ticks display 百, 千, and 万. | Generate Excel workbooks for the Japanese market with consistent chart unit localization across multiple sheets. | Reuse a single ChartJapaneseSettings class in different .NET projects to enforce Japanese chart conventions.
// AI Prompts: Demonstrate how to implement a ChartJapaneseSettings class that overrides GetAxisUnitName for Japanese units and attach it to workbook.Settings.GlobalizationSettings.ChartSettings before saving. | Show how to apply the same custom ChartGlobalizationSettings to several charts in one workbook using Aspose.Cells for .NET. | Explain how to extend ChartJapaneseSettings to support additional DisplayUnitType values such as Millions while keeping existing Japanese unit names.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to subclass ChartGlobalizationSettings, override GetAxisUnitName to return Japanese characters for hundreds, thousands and ten‑thousands, and assign the custom ChartJapaneseSettings to workbook.Settings.GlobalizationSettings.ChartSettings before saving the workbook as an Excel file.
class ChartJapaneseSettings : ChartGlobalizationSettings
{
    // Provide Japanese specific axis unit names
    public override string GetAxisUnitName(DisplayUnitType type)
    {
        switch (type)
        {
            case DisplayUnitType.Hundreds:
                return "百";
            case DisplayUnitType.Thousands:
                return "千";
            case DisplayUnitType.TenThousands:
                return "万";
            default:
                return base.GetAxisUnitName(type);
        }
    }
}

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";
        chart.Title.Text = "サンプルチャート";

        // Assign Japanese globalization settings to the workbook's chart settings
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = new ChartJapaneseSettings()
        };

        // Export the workbook
        workbook.Save("ChartJapaneseSettingsDemo.xlsx");
    }
}
