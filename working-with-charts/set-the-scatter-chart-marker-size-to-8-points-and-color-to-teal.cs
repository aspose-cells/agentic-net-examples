// Title: How to set a teal 8‑point marker for a scatter chart series using Aspose.Cells in C#
// AI Prompts: Generate C# code with Aspose.Cells that creates a scatter chart and configures the series markers to be circular, 8 points in size, and filled with teal. | Update an existing Aspose.Cells workbook to change the scatter series marker style to a custom teal color and a specific point size.
// Common Searches: Aspose.Cells C# set scatter chart marker size to 8 points | change scatter series marker color to teal using Aspose.Cells | C# Aspose.Cells customize marker style for scatter chart | set custom marker size and fill color for Aspose.Cells chart series | example of teal markers in Aspose.Cells scatter plot
// Tags: scatter chart marker customization Aspose.Cells | adjust marker dimensions Aspose.Cells C# | custom color fill for chart markers Aspose.Cells | chart series marker style C# Aspose.Cells | Excel chart marker formatting Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills columns A and B with X/Y data, adds a scatter chart, links the data to the series, and then sets the series marker to a circular shape, 8‑point size, and teal fill before saving the file as ScatterChartWithCustomMarker.xlsx.
class ScatterChartMarkerDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for X and Y values
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[i, 0].PutValue(i - 1);          // X values: 1,2,3,4,5
            sheet.Cells[i, 1].PutValue((i - 1) * 10);   // Y values: 10,20,30,40,50
        }

        // Add a scatter chart
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 7, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the series (X values are taken from column A, Y from column B)
        chart.NSeries.Add("B2:B6", true);
        chart.NSeries[0].XValues = "A2:A6";

        // Configure marker: size = 8 points, color = teal
        Series series = chart.NSeries[0];
        series.Marker.MarkerStyle = ChartMarkerType.Circle; // optional, ensures markers are visible
        series.Marker.MarkerSize = 8;                        // size in points
        series.Marker.ForegroundColor = Color.Teal;         // marker fill color
        series.Marker.ForegroundColorSetType = FormattingType.Custom; // apply custom color

        // Save the workbook
        workbook.Save("ScatterChartWithCustomMarker.xlsx");
    }
}
