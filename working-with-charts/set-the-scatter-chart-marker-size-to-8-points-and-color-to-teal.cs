using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ScatterChartMarkerDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for X and Y values
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[i, 0].PutValue(i - 1);          // X values: 1,2,3,4,5
            sheet.Cells[i, 1].PutValue((i - 1) * 2);    // Y values: 2,4,6,8,10
        }

        // Add a scatter chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 7, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart (Y values) and X values separately
        chart.NSeries.Add("B2:B6", true);
        chart.NSeries[0].XValues = "A2:A6";

        // Configure marker: size = 8 points, color = teal
        Series series = chart.NSeries[0];
        series.Marker.MarkerStyle = ChartMarkerType.Circle; // optional style
        series.Marker.MarkerSize = 8;                        // size in points
        series.Marker.ForegroundColor = Color.Teal;         // marker color

        // Save the workbook
        workbook.Save("ScatterChartWithCustomMarker.xlsx");
    }
}