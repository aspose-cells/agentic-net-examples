using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMarkerDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the line chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["B5"].PutValue(25);

            // Add a line chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories (X‑axis)
            chart.NSeries.Add("B2:B5", true);          // Values
            chart.NSeries.CategoryData = "A2:A5";      // Categories

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Enable data markers and set the marker shape to triangle
            series.Marker.MarkerStyle = ChartMarkerType.Triangle;
            series.Marker.MarkerSize = 12;               // Optional: adjust marker size
            series.Marker.ForegroundColor = Color.Blue; // Optional: set marker foreground color
            series.Marker.BackgroundColor = Color.LightYellow; // Optional: set marker background color

            // Save the workbook to an Excel file
            workbook.Save("LineChartWithTriangleMarkers.xlsx");
        }
    }
}