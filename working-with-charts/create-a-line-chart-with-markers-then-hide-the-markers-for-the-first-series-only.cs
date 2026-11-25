// Title: Create a line chart with data markers and hide the markers for the first series using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide markers for only the first series in a line chart | set marker style to none for a specific series in an Aspose.Cells line chart | add line chart with data markers and customize second series markers in C# | how to configure chart series markers in an Aspose.Cells workbook | generate Excel line chart with markers and hide them for one series using Aspose.Cells
// Tags: line chart hide first series markers Aspose.Cells | set chart series marker style none C# | add line chart with data markers Aspose.Cells | customize second series marker color and size Aspose.Cells | save workbook as xlsx Aspose.Cells | configure chart title and axis titles Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsLineChartExample
{
    // The example creates a new workbook, populates sample data, adds a LineWithDataMarkers chart, assigns category data, adds two series, hides markers for the first series by setting MarkerStyle to None, customizes the second series markers (circle, size 8, blue), sets chart and axis titles, and saves the workbook as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A: Categories (X‑axis)
            // Column B: Series 1 values
            // Column C: Series 2 values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");

            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);      // Series 1 data
                sheet.Cells[$"C{i}"].PutValue(i * 12 + 5); // Series 2 data
            }

            // Add a line chart with data markers
            int chartIndex = sheet.Charts.Add(ChartType.LineWithDataMarkers, 8, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the category (X‑axis) data for the chart
            chart.NSeries.CategoryData = "A2:A6";

            // Add the first series (Series 1)
            int series1Idx = chart.NSeries.Add("B2:B6", false);
            // Hide markers for the first series
            chart.NSeries[series1Idx].Marker.MarkerStyle = ChartMarkerType.None;

            // Add the second series (Series 2)
            int series2Idx = chart.NSeries.Add("C2:C6", false);
            // Keep default markers (automatic) or customize if desired
            chart.NSeries[series2Idx].Marker.MarkerStyle = ChartMarkerType.Circle;
            chart.NSeries[series2Idx].Marker.MarkerSize = 8;
            chart.NSeries[series2Idx].Marker.ForegroundColor = Color.Blue;

            // Optional: set chart title and axis titles
            chart.Title.Text = "Line Chart – First Series Markers Hidden";
            chart.CategoryAxis.Title.Text = "Category";
            chart.ValueAxis.Title.Text = "Value";

            // Save the workbook
            string outputPath = "LineChartHideFirstSeriesMarkers.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
