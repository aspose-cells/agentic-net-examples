// Title: C# – Create a Line Chart with Selective Markers (hide first series) using Aspose.Cells
// Description: This example builds a workbook, adds category data and two series, inserts a LineWithDataMarkers chart, hides markers for the first series, customizes markers for the second series, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells line chart C# | hide markers first series | custom chart markers Aspose.Cells | .NET chart series marker style | LineWithDataMarkers Aspose | Excel chart marker customization
// Common Searches: Aspose.Cells hide markers for a specific series | C# line chart marker style none Aspose | how to customize chart markers in Aspose.Cells | selective marker visibility in Excel chart .NET | LineWithDataMarkers example Aspose
// Developer Intent: Generate a line chart with data markers, suppress markers for the first series, and apply custom marker styling to the second series using Aspose.Cells for .NET.
// Use Cases: Display a baseline trend line without markers while highlighting a comparison series with colored circles. | Create a performance dashboard where only key data points are emphasized with custom markers. | Export Excel reports that require different marker visibility per series to match corporate presentation standards.
// AI Prompts: Show C# code to hide markers for the first series of a line chart while keeping markers for other series in Aspose.Cells. | Provide an Aspose.Cells example that customizes marker shape, size, and colors for a specific series. | Explain how to add a LineWithDataMarkers chart, set the first series marker style to None, and style the second series markers programmatically.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example builds a workbook, adds category data and two series, inserts a LineWithDataMarkers chart, hides markers for the first series, customizes markers for the second series, and saves the file as an Excel workbook.
class LineChartWithSelectiveMarkers
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for two series
        // Category (X axis)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["A5"].PutValue("Apr");

        // First series values
        sheet.Cells["B1"].PutValue("Series 1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(40);

        // Second series values
        sheet.Cells["C1"].PutValue("Series 2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);
        sheet.Cells["C5"].PutValue(45);

        // Add a line chart that includes data markers by default
        int chartIndex = sheet.Charts.Add(ChartType.LineWithDataMarkers, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set data ranges for the two series
        chart.NSeries.Add("B2:B5", true); // Series 1
        chart.NSeries.Add("C2:C5", true); // Series 2
        chart.NSeries.CategoryData = "A2:A5";

        // Customize markers for the second series (keep default style or set explicitly)
        Series secondSeries = chart.NSeries[1];
        secondSeries.Marker.MarkerStyle = ChartMarkerType.Circle;
        secondSeries.Marker.MarkerSize = 8;
        secondSeries.Marker.ForegroundColor = Color.Blue;
        secondSeries.Marker.BackgroundColor = Color.LightBlue;

        // Hide markers for the first series by setting marker style to None
        Series firstSeries = chart.NSeries[0];
        firstSeries.Marker.MarkerStyle = ChartMarkerType.None;

        // Save the workbook
        workbook.Save("LineChartSelectiveMarkers.xlsx");
    }
}
