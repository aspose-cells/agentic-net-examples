// Title: Aspose.Cells for .NET – Set Scatter Chart Marker Size to 8 pt and Color to Teal (C#)
// Description: Creates a workbook, fills columns A and B with X/Y values, adds a scatter chart, and configures the first series marker to a circular shape, 8‑point size, teal foreground color using custom formatting, then saves the file as ScatterMarkerTeal.xlsx.
// Keywords: Aspose.Cells | C# | scatter chart | marker size | marker color | teal | ChartMarkerType.Circle | FormattingType.Custom | Excel chart customization | set marker size .NET
// Common Searches: Aspose.Cells set scatter chart marker size | C# change scatter plot marker color to teal | how to customize marker style in Aspose.Cells chart | set marker size 8 points Aspose.Cells | scatter chart marker customization .NET
// Developer Intent: Set the marker size to 8 points and apply a teal color to a scatter chart series using Aspose.Cells in C#.
// Use Cases: Produce Excel reports where scatter points are highlighted with a specific size and color for better visual emphasis. | Standardize chart appearance across multiple workbooks by programmatically defining marker properties. | Create presentation‑ready scatter charts with consistent styling for dashboards or data analysis tools.
// AI Prompts: Generate C# code with Aspose.Cells to set scatter chart markers to 8 pt size and teal color. | Explain how to apply a custom teal foreground color to multiple series in an Aspose.Cells scatter chart. | Show the steps to ensure FormattingType.Custom forces the marker color change in Aspose.Cells.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills columns A and B with X/Y values, adds a scatter chart, and configures the first series marker to a circular shape, 8‑point size, teal foreground color using custom formatting, then saves the file as ScatterMarkerTeal.xlsx.
class SetScatterMarker
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample X and Y data for the scatter chart
        worksheet.Cells["A1"].PutValue("X");
        worksheet.Cells["B1"].PutValue("Y");
        for (int i = 1; i <= 5; i++)
        {
            worksheet.Cells[i, 0].PutValue(i);          // X values in column A
            worksheet.Cells[i, 1].PutValue(i * 2);      // Y values in column B
        }

        // Add a scatter chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series (Y values) and X values separately
        chart.NSeries.Add("B2:B6", true);
        chart.NSeries[0].XValues = "A2:A6";

        // Access the first series and configure its marker
        Series series = chart.NSeries[0];
        series.Marker.MarkerStyle = ChartMarkerType.Circle;   // Use circular markers
        series.Marker.MarkerSize = 8;                         // Size in points
        series.Marker.ForegroundColor = Color.Teal;          // Marker color
        series.Marker.ForegroundColorSetType = FormattingType.Custom; // Ensure custom color is applied

        // Save the workbook with the configured scatter chart
        workbook.Save("ScatterMarkerTeal.xlsx");
    }
}
