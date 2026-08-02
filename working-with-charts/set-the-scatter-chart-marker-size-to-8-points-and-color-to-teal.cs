using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the scatter chart (X and Y values)
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 1; i <= 5; i++)
        {
            sheet.Cells[i, 0].PutValue(i);        // X values in column A
            sheet.Cells[i, 1].PutValue(i * 2);    // Y values in column B
        }

        // Add a scatter chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series (Y values) and X values separately
        chart.NSeries.Add("B2:B6", true);
        chart.NSeries[0].XValues = "A2:A6";

        // Configure the marker: size = 8 points, color = teal, style = circle
        chart.NSeries[0].Marker.MarkerSize = 8;               // size in points
        chart.NSeries[0].Marker.ForegroundColor = Color.Teal; // marker color
        chart.NSeries[0].Marker.MarkerStyle = ChartMarkerType.Circle;

        // Save the workbook with the configured scatter chart
        workbook.Save("ScatterChartMarker.xlsx");
    }
}