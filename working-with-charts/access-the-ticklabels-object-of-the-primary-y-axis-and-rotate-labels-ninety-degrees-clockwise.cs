using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class RotateYAxisTickLabels
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category 1");
        sheet.Cells["A2"].PutValue("Category 2");
        sheet.Cells["A3"].PutValue("Category 3");
        sheet.Cells["B1"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["B3"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source
        chart.NSeries.Add("B1:B3", true);
        chart.NSeries.CategoryData = "A1:A3";

        // Access the primary Y axis (value axis) tick labels and rotate them 90 degrees clockwise
        chart.ValueAxis.TickLabels.RotationAngle = 90;

        // Save the workbook to a file
        workbook.Save("YAxisTickLabelsRotated.xlsx");
    }
}