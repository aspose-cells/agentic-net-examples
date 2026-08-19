// Title: Add Triangle Markers to a Line Series with Aspose.Cells for .NET
// Description: Shows how to build a workbook, fill it with sample data, insert a Line chart, enable data markers, set the marker style to a triangle, adjust the marker size, and save the workbook as an Excel file using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# chart marker | line chart triangle marker | ChartMarkerType.Triangle | enable data markers | ChartType.Line | Excel export .NET | custom chart markers | Aspose.Cells example
// Common Searches: Aspose.Cells set triangle marker on line chart | how to enable data markers in Aspose.Cells C# | change marker shape to triangle in Excel chart programmatically | customize line series markers with Aspose.Cells | C# Aspose.Cells chart marker size
// Developer Intent: Add triangle‑shaped data markers to a line‑chart series programmatically.
// Use Cases: Highlight key sales figures on a trend line with distinct triangle markers for presentations. | Create a performance dashboard where each measurement point is emphasized by a custom marker shape. | Generate an Excel report that includes a line chart with triangle markers to improve data readability.
// AI Prompts: Write C# code using Aspose.Cells to add a line chart and set its series markers to a triangle with a specific size. | Provide an Aspose.Cells example that applies different marker shapes to multiple series, including triangles, squares, and circles. | Explain the steps to enable data markers and customize their appearance (shape, size, color) for a line chart in Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to build a workbook, fill it with sample data, insert a Line chart, enable data markers, set the marker style to a triangle, adjust the marker size, and save the workbook as an Excel file using Aspose.Cells in C#.
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

        // Define the data series and category labels
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data markers and set the marker shape to triangle
        Series series = chart.NSeries[0];
        series.Marker.MarkerStyle = ChartMarkerType.Triangle;
        series.Marker.MarkerSize = 10; // optional: set marker size

        // Save the workbook with the chart
        workbook.Save("LineChartWithTriangleMarkers.xlsx");
    }
}
