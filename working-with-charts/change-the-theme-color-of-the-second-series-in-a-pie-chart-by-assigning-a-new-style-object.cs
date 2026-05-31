using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class ChangeSecondSeriesThemeColor
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data for a pie chart with two series
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");

        // First series values
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(50);

        // Second series values
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(40);
        sheet.Cells["C3"].PutValue(35);
        sheet.Cells["C4"].PutValue(25);

        // Add a pie chart
        int chartIdx = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIdx];

        // Add both series to the chart
        chart.NSeries.Add("B2:B4", true); // first series
        chart.NSeries.Add("C2:C4", true); // second series
        chart.NSeries.CategoryData = "A2:A4";

        // Create a new Style object and set its foreground color (theme color)
        Style newStyle = workbook.CreateStyle();
        // Example: use Accent2 theme color (you can choose any ThemeColorType)
        newStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent2, 0.0);
        // For the purpose of applying to the series, extract the actual Color from the theme
        // Here we simply use a concrete color; in a real scenario you might resolve the theme color.
        Color themeColor = Color.FromArgb(192, 80, 77); // a distinct color for demonstration
        newStyle.ForegroundColor = themeColor;
        newStyle.Pattern = BackgroundType.Solid;

        // Assign the style's foreground color to the second series area
        // This changes the visual color of the second series in the pie chart
        chart.NSeries[1].Area.ForegroundColor = newStyle.ForegroundColor;

        // Save the workbook
        workbook.Save("PieChart_SecondSeriesThemeColor.xlsx");
    }
}