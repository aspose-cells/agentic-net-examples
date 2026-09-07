// Title: Add blue triangle data markers to a line series in an Aspose.Cells .NET workbook (C#)
// AI Prompts: Generate C# code using Aspose.Cells to create a line chart and set the first series markers to blue triangles with a size of 10. | Show how to enable data markers on a line series and customize marker style, size, and color with Aspose.Cells for .NET. | Write a snippet that adds a line chart, binds category and value ranges, and configures the series to use triangle markers.
// Common Searches: Aspose.Cells C# line chart triangle marker shape example | how to set marker color and size for line series in Aspose.Cells .NET | enable data markers on line chart using Aspose.Cells API | C# Aspose.Cells change line chart series marker to triangle | Aspose.Cells line chart custom marker style tutorial
// Tags: Aspose.Cells line chart marker customization C# | set triangle marker style Aspose.Cells series | configure data marker size and color Aspose.Cells | line series marker shape Aspose.Cells .NET | Excel chart triangle markers using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Demonstrates creating a workbook, adding a line chart, binding category and value ranges, and configuring the first series to display blue triangle markers of size 10 before saving as LineChart_TriangleMarkers.xlsx.
class EnableTriangleMarkers
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

        // Define the data series for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Access the first series and configure its markers
        Series series = chart.NSeries[0];
        series.Marker.MarkerStyle = ChartMarkerType.Triangle; // Set marker shape to triangle
        series.Marker.MarkerSize = 10;                         // Optional: set marker size
        series.Marker.ForegroundColor = Color.Blue;           // Optional: set marker color

        // Save the workbook with the chart
        workbook.Save("LineChart_TriangleMarkers.xlsx");
    }
}
