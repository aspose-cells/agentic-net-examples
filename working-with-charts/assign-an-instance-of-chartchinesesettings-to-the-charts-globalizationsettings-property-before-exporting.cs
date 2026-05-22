using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Add sample data for the chart
        ws.Cells["A1"].PutValue("类别");
        ws.Cells["A2"].PutValue("第一组");
        ws.Cells["A3"].PutValue("第二组");
        ws.Cells["B1"].PutValue("数值");
        ws.Cells["B2"].PutValue(120);
        ws.Cells["B3"].PutValue(250);

        // Create a column chart
        int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = ws.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Apply custom Chinese globalization settings to the chart via workbook settings
        wb.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = new ChartChineseSettings()
        };

        // Export the workbook (saving to a file)
        wb.Save("ChartWithChineseSettings.xlsx");
    }

    // Custom globalization settings for Chinese language
    class ChartChineseSettings : ChartGlobalizationSettings
    {
        // Override axis unit names to Chinese characters
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

        // Override chart title name
        public override string GetChartTitleName()
        {
            return "图表标题";
        }

        // Additional overrides can be added here as needed
    }
}