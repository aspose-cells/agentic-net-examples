using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ChartJapaneseSettings : ChartGlobalizationSettings
{
    // Override axis unit names with Japanese equivalents
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

    // Additional overrides can be added here if needed
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
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Assign custom Japanese globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = new ChartJapaneseSettings()
        };

        // Export the workbook (e.g., to XLSX)
        workbook.Save("ChartJapanese.xlsx");
    }
}