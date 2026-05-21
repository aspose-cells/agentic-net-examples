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

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the first series and enable data markers with a triangle shape
        Series series = chart.NSeries[0];
        series.Marker.MarkerStyle = ChartMarkerType.Triangle; // Set marker shape to triangle
        series.Marker.MarkerSize = 10;                         // Optional: set marker size
        series.Marker.ForegroundColor = Color.Blue;           // Optional: set marker foreground color
        series.Marker.BackgroundColor = Color.LightYellow;   // Optional: set marker background color

        // Save the workbook to a file
        workbook.Save("LineChartWithTriangleMarkers.xlsx");
    }
}