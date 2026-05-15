using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

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
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
        }

        // Add a column chart and set its data range in a single call
        // Parameters: ChartType, dataRange, isVertical, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.Column, "A1:B5", true, 7, 1, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Customize the legend: place it at the bottom and ensure it does not overlay the chart
        chart.Legend.Position = LegendPositionType.Bottom;
        chart.Legend.IsOverLay = false;

        // Save the workbook with the chart
        workbook.Save("ChartWithLegend.xlsx", SaveFormat.Xlsx);
    }
}