using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetLegendBottomLeft
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set legend position to Bottom (bottom left)
        Legend legend = chart.Legend;
        legend.Position = LegendPositionType.Bottom;

        // Align legend with the left margin of the chart
        // XRatioToChart is a value between 0.0 (left edge) and 1.0 (right edge)
        legend.XRatioToChart = 0.0; // left-aligned

        // Save the workbook
        workbook.Save("LegendBottomLeft.xlsx");
    }
}