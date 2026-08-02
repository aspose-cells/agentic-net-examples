// Title: C# Aspose.Cells Example: Line Chart with Markers – Hide Markers on First Series
// Description: Creates a workbook, adds category and two data series, inserts a LineWithDataMarkers chart, shows circular markers only on the second series, hides markers on the first series, and saves the file as an XLSX document.
// Keywords: Aspose.Cells line chart C# | LineWithDataMarkers | chart markers Aspose.Cells | hide series markers | ChartMarkerType.None | .NET Excel chart example | selective marker visibility | Aspose.Cells API sample
// Common Searches: Aspose.Cells hide markers for a specific series | C# line chart with markers Aspose.Cells | set ChartMarkerType.None Aspose.Cells .NET | example of selective markers in Excel chart | how to display markers only on one series using Aspose.Cells
// Developer Intent: Generate a line chart where markers appear only on the second data series while the first series is rendered without markers.
// Use Cases: Highlight a key product’s trend in a sales dashboard by adding markers only to its line series. | Create a clean performance chart that shows raw data points for one metric and a plain line for another. | Produce an Excel report that emphasizes specific data series through visible markers while keeping other series uncluttered.
// AI Prompts: Write C# code with Aspose.Cells to add a LineWithDataMarkers chart and hide markers on the first series. | Show how to apply different ChartMarkerType values to multiple series in an Aspose.Cells line chart. | Explain the effect of setting ChartMarkerType.None on a series in an Aspose.Cells-generated Excel chart.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds category and two data series, inserts a LineWithDataMarkers chart, shows circular markers only on the second series, hides markers on the first series, and saves the file as an XLSX document.
class LineChartWithSelectiveMarkers
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Series 1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series 2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Add a line chart that supports data markers
        int chartIndex = sheet.Charts.Add(ChartType.LineWithDataMarkers, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data ranges for the two series
        chart.NSeries.Add("B2:B4", true); // Series 1
        chart.NSeries.Add("C2:C4", true); // Series 2
        chart.NSeries.CategoryData = "A2:A4";

        // Configure markers for the second series (visible)
        chart.NSeries[1].Marker.MarkerStyle = ChartMarkerType.Circle;
        chart.NSeries[1].Marker.MarkerSize = 8;
        chart.NSeries[1].Marker.ForegroundColor = Color.Blue;

        // Hide markers for the first series
        chart.NSeries[0].Marker.MarkerStyle = ChartMarkerType.None;

        // Save the workbook
        workbook.Save("LineChartSelectiveMarkers.xlsx", SaveFormat.Xlsx);
    }
}
