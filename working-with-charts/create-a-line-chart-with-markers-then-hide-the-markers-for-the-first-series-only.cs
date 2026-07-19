// Title: Aspose.Cells .NET: Create a Line Chart with Selective Markers (Hide First Series)
// Description: This example shows how to generate an XLSX workbook, populate it with category and two data series, add a LineWithDataMarkers chart, apply circular markers to the second series, set the first series marker style to none, and save the file. It demonstrates precise control over marker visibility for individual chart series using Aspose.Cells for C#.
// Keywords: Aspose.Cells line chart | C# line chart markers | hide series markers Aspose.Cells | ChartMarkerType.None | LineWithDataMarkers example | .NET Excel chart customization | selective markers Aspose.Cells | Excel line chart styling C# | Aspose.Cells chart series formatting | marker visibility Excel chart
// Common Searches: Aspose.Cells hide markers for one series | C# line chart with data markers Aspose.Cells | set marker style none Aspose.Cells chart | selective marker visibility in Excel chart .NET | how to customize markers per series Aspose.Cells
// Developer Intent: Create a line chart where only the second series displays markers while the first series has markers disabled.
// Use Cases: Display a forecast line without markers and a actual line with colored markers for clearer comparison. | Generate a sales trend chart that highlights only the secondary series data points to reduce visual clutter. | Produce a financial report where baseline values are shown as a smooth line and key performance indicators are emphasized with markers.
// AI Prompts: Write C# code using Aspose.Cells to add a LineWithDataMarkers chart, set circular markers for series 2, and hide markers for series 1. | Explain how to apply ChartMarkerType.None to a specific series in an Aspose.Cells line chart and when this technique is useful. | Provide step‑by‑step instructions to customize marker colors for one series while disabling markers for another in an Aspose.Cells .NET workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example shows how to generate an XLSX workbook, populate it with category and two data series, add a LineWithDataMarkers chart, apply circular markers to the second series, set the first series marker style to none, and save the file. It demonstrates precise control over marker visibility for individual chart series using Aspose.Cells for C#.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a line chart with data markers
                int chartIdx = sheet.Charts.Add(ChartType.LineWithDataMarkers, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Set data ranges for the two series
                chart.NSeries.Add("B2:B4", true); // Series 1 values
                chart.NSeries.Add("C2:C4", true); // Series 2 values
                chart.NSeries.CategoryData = "A2:A4";

                // Configure markers for the second series (visible)
                Series series2 = chart.NSeries[1];
                series2.Marker.MarkerStyle = ChartMarkerType.Circle;
                series2.Marker.MarkerSize = 8;
                series2.Marker.ForegroundColor = Color.Blue;
                series2.Marker.BackgroundColor = Color.LightBlue;

                // Hide markers for the first series
                Series series1 = chart.NSeries[0];
                series1.Marker.MarkerStyle = ChartMarkerType.None;

                // Define output file path
                string outputPath = "LineChartWithSelectiveMarkers.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
