using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ComboChartMarkerDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Column A: Categories
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["A5"].PutValue("Apr");

        // Column B: Sales (column series)
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);
        sheet.Cells["B5"].PutValue(200);

        // Column C: Profit (line series)
        sheet.Cells["C1"].PutValue("Profit");
        sheet.Cells["C2"].PutValue(30);
        sheet.Cells["C3"].PutValue(45);
        sheet.Cells["C4"].PutValue(50);
        sheet.Cells["C5"].PutValue(70);

        // Add a combo chart (initially a column chart)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add the first series (Sales) – stays as column
        chart.NSeries.Add("B2:B5", true);
        // Add the second series (Profit) – will be changed to line
        chart.NSeries.Add("C2:C5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Convert the second series to a line series (creates a combo chart)
        Series lineSeries = chart.NSeries[1];
        lineSeries.Type = ChartType.Line;

        // Customize marker style for the entire line series
        lineSeries.Marker.MarkerStyle = ChartMarkerType.Circle;
        lineSeries.Marker.MarkerSize = 10;               // size in points
        lineSeries.Marker.ForegroundColor = Color.Red;   // marker border color
        lineSeries.Marker.BackgroundColor = Color.Yellow; // marker fill color

        // Highlight specific key data points (e.g., March and April)
        // March (index 2)
        ChartPoint marchPoint = lineSeries.Points[2];
        marchPoint.Marker.MarkerStyle = ChartMarkerType.Square;
        marchPoint.Marker.MarkerSize = 14;
        marchPoint.Marker.ForegroundColor = Color.Blue;
        marchPoint.Marker.BackgroundColor = Color.LightGreen;

        // April (index 3)
        ChartPoint aprilPoint = lineSeries.Points[3];
        aprilPoint.Marker.MarkerStyle = ChartMarkerType.Diamond;
        aprilPoint.Marker.MarkerSize = 14;
        aprilPoint.Marker.ForegroundColor = Color.Purple;
        aprilPoint.Marker.BackgroundColor = Color.Orange;

        // Save the workbook with the customized combo chart
        workbook.Save("ComboChartWithCustomMarkers.xlsx");
    }
}