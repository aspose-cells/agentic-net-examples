// Title: Apply Japanese ChartGlobalizationSettings in Aspose.Cells (C#) before saving
// Description: Creates a workbook, adds sample data, builds a column chart, defines a ChartJapaneseSettings class that overrides chart titles, axis labels, legend entries, and unit names, assigns this class to Workbook.Settings.GlobalizationSettings.ChartSettings, and saves the file so the chart displays Japanese text.
// Keywords: Aspose.Cells | C# chart globalization | ChartJapaneseSettings | Japanese chart localization | ChartGlobalizationSettings example | Excel chart Japanese titles | Aspose.Cells Japanese axis units | globalization settings Aspose.Cells
// Common Searches: Aspose.Cells set Japanese language for charts | How to use ChartGlobalizationSettings in .NET | Assign custom ChartJapaneseSettings to workbook | Localize Excel chart titles with Aspose.Cells | Override GetAxisUnitName for Japanese units
// Developer Intent: Localize chart elements to Japanese by assigning a ChartJapaneseSettings object to the workbook’s globalization settings before export.
// Use Cases: Produce sales charts for Japanese stakeholders with native titles, legends, and axis units. | Create a reusable localization class that automatically translates chart text for multiple Excel reports. | Display custom unit symbols (百, 千, 万) on chart axes in Japanese‑language workbooks.
// AI Prompts: Generate a C# example that implements ChartGlobalizationSettings for French and applies it to an Aspose.Cells workbook. | Show code to assign different ChartJapaneseSettings instances to several charts in the same workbook. | Explain how to test that overridden GetAxisUnitName values appear correctly after saving the Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, builds a column chart, defines a ChartJapaneseSettings class that overrides chart titles, axis labels, legend entries, and unit names, assigns this class to Workbook.Settings.GlobalizationSettings.ChartSettings, and saves the file so the chart displays Japanese text.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate sample data for the chart
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["A2"].PutValue("Jan");
        ws.Cells["A3"].PutValue("Feb");
        ws.Cells["B1"].PutValue("Sales");
        ws.Cells["B2"].PutValue(100);
        ws.Cells["B3"].PutValue(150);

        // Add a column chart to the worksheet
        int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = ws.Charts[chartIdx];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";
        chart.Title.Text = "売上";

        // Assign Japanese globalization settings to the workbook (affects the chart)
        wb.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = new ChartJapaneseSettings()
        };

        // Export the workbook
        wb.Save("ChartJapaneseSettingsDemo.xlsx");
    }
}

// Custom Japanese chart globalization settings
class ChartJapaneseSettings : ChartGlobalizationSettings
{
    public override string GetChartTitleName()
    {
        return "チャートタイトル";
    }

    public override string GetAxisTitleName()
    {
        return "軸タイトル";
    }

    public override string GetLegendIncreaseName()
    {
        return "増加";
    }

    public override string GetLegendDecreaseName()
    {
        return "減少";
    }

    public override string GetOtherName()
    {
        return "その他";
    }

    public override string GetSeriesName()
    {
        return "シリーズ";
    }

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
