using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ChartLegendBottomExample
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Position the legend at the bottom of the chart
        chart.Legend.Position = LegendPositionType.Bottom;

        // Optional: adjust legend size if needed
        chart.Legend.IsAutomaticSize = false;
        chart.Legend.Width = 400;
        chart.Legend.Height = 50;

        // Save the workbook (lifecycle: save)
        workbook.Save("ChartWithBottomLegend.xlsx");
    }
}