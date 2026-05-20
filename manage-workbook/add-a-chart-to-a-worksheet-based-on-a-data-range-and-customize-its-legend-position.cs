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

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[$"A{i}"].PutValue("Item " + (i - 1));
            sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
        }

        // Add a column chart (rows 8‑20, columns 1‑8)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 19, 7);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart (A1:B6) and plot by column
        chart.SetChartDataRange("A1:B6", true);

        // Customize the legend: place it at the bottom of the chart
        chart.Legend.Position = LegendPositionType.Bottom;
        chart.ShowLegend = true; // Ensure the legend is visible

        // Save the workbook with the chart
        workbook.Save("ChartWithBottomLegend.xlsx", SaveFormat.Xlsx);
    }
}