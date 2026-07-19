// Title: Aspose.Cells for .NET – C# Example: Set Triangle Markers on a Line Chart Series
// Description: The sample creates a workbook, fills category and value cells, adds a line chart, links the series to the data range, turns on data markers, changes the marker style to a triangle, optionally sets the marker size, and writes the result to an Excel file.
// Keywords: Aspose.Cells | C# | .NET chart | line chart markers | triangle marker | ChartMarkerType | marker size | Excel automation | custom chart series | data point markers
// Common Searches: Aspose.Cells set triangle marker on line chart | C# enable data markers in Aspose.Cells chart | how to change marker shape in Aspose.Cells line series | customize marker size Aspose.Cells .NET | line chart marker style example Aspose
// Developer Intent: Apply triangular markers to a line‑chart series with Aspose.Cells for .NET.
// Use Cases: Highlight each point in a sales‑trend line chart with a distinct triangle for quick visual scanning. | Produce a financial dashboard where key performance values are emphasized using custom‑sized triangle markers. | Generate a corporate‑branded report that matches the company’s visual guidelines by applying triangle markers to all line series.
// AI Prompts: Generate C# code that adds a line chart to a workbook and sets the series markers to triangles of a given size using Aspose.Cells. | Show how to enable markers for multiple series in an Aspose.Cells chart and assign different shapes, including a triangle, to each series. | Explain the steps to programmatically modify marker style and dimensions for a line series in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills category and value cells, adds a line chart, links the series to the data range, turns on data markers, changes the marker style to a triangle, optionally sets the marker size, and writes the result to an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
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

        // Access the first series and enable data markers
        Series series = chart.NSeries[0];
        series.Marker.MarkerStyle = ChartMarkerType.Triangle; // Set marker shape to triangle
        series.Marker.MarkerSize = 10; // Optional: set marker size

        // Save the workbook to a file
        workbook.Save("LineChartWithTriangleMarkers.xlsx");
    }
}
