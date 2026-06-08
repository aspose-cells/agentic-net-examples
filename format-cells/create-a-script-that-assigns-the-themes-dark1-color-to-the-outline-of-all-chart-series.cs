using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class AssignDark1Outline
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["C1"].PutValue("Series2");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells["A" + i].PutValue("Cat " + (i - 1));
            sheet.Cells["B" + i].PutValue(i * 10);
            sheet.Cells["C" + i].PutValue(i * 15);
        }

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("B1:C6", true);
        chart.NSeries.CategoryData = "A2:A6";

        // Assign the theme's Dark1 color (Background1) to the outline of all series
        foreach (Series series in chart.NSeries)
        {
            // Set the border's theme color to Background1 with no tint (full opacity)
            series.Border.ThemeColor = new ThemeColor(ThemeColorType.Background1, 0);
            series.Border.IsVisible = true; // Ensure the border is visible
        }

        // Save the workbook
        workbook.Save("ChartSeriesDark1Outline.xlsx");
    }
}