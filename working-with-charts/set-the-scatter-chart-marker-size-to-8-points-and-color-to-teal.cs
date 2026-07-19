// Title: Aspose.Cells .NET – Set Scatter Chart Marker Size to 8 pt and Color to Teal
// Description: Creates a workbook, adds X/Y data, inserts a scatter chart, binds the series, and customizes the marker to a teal circle with an 8‑point size before saving the file.
// Keywords: Aspose.Cells scatter chart marker | C# set marker size | teal chart marker color | scatter plot customization .NET | Excel chart marker style
// Common Searches: Aspose.Cells change scatter marker size C# | set teal color for chart markers Aspose.Cells | C# scatter chart marker style example | how to format scatter plot markers in .NET
// Developer Intent: Apply a teal circular marker of 8 pt to a scatter chart series using Aspose.Cells for .NET.
// Use Cases: Highlight key data points in automated Excel reports with a consistent teal marker. | Standardize visual styling of scatter charts across multiple generated workbooks. | Create visually distinct series in dashboards that require precise marker sizing and coloring.
// AI Prompts: Generate C# code with Aspose.Cells to set a scatter chart marker to 8 pt teal color. | Explain step‑by‑step how to customize marker style, size, and color for a scatter chart in Aspose.Cells .NET. | Show how to apply the same teal 8‑point marker to all series in a scatter chart using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

// Creates a workbook, adds X/Y data, inserts a scatter chart, binds the series, and customizes the marker to a teal circle with an 8‑point size before saving the file.
class SetScatterMarker
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the scatter chart
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[i, 0].PutValue(i - 1);          // X values
            sheet.Cells[i, 1].PutValue((i - 1) * 2);    // Y values
        }

        // Add a scatter chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add a series and bind X and Y values
        chart.NSeries.Add("B2:B6", true);
        chart.NSeries[0].XValues = "A2:A6";

        // Configure marker: size 8 points, teal color
        Series series = chart.NSeries[0];
        series.Marker.MarkerStyle = ChartMarkerType.Circle;
        series.Marker.MarkerSize = 8;               // size in points
        series.Marker.ForegroundColor = Color.Teal; // marker color

        // Save the workbook
        workbook.Save("ScatterMarkerTeal.xlsx");
    }
}
