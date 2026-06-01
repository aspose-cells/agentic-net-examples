using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 5; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Apply a vertical gradient fill (white at top, gray at bottom) to the plot area
        FillFormat plotAreaFill = chart.PlotArea.Area.FillFormat;
        plotAreaFill.SetTwoColorGradient(Color.White, Color.Gray, GradientStyleType.Vertical, 1);

        // Save the workbook
        workbook.Save("ChartPlotAreaGradient.xlsx");
    }
}