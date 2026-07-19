// Title: Aspose.Cells C# – Customize Line Series Markers and Highlight Points in a Combo Chart
// Description: Creates an Excel workbook with sales and target data, adds a combo chart (column + line), applies circular red‑bordered yellow markers to the entire line series, and emphasizes a chosen point with a larger blue‑bordered green square before saving the file.
// Keywords: Aspose.Cells | C# chart customization | combo chart markers | line series marker style | highlight chart point | ChartMarkerType | ChartPoint formatting | Excel visual emphasis | .NET Excel chart API
// Common Searches: Aspose.Cells change line series marker style C# | How to highlight a specific point in an Aspose.Cells chart | Set marker color and size for combo chart in .NET | Customize markers for line series in Aspose.Cells | Mark a key data point in an Excel combo chart using C#
// Developer Intent: Apply custom marker formatting to a line series in a combo chart and draw attention to a particular data point.
// Use Cases: Visually separate target values from column data by using distinct markers on the line series. | Draw audience focus to a milestone or outlier by giving one point a larger, uniquely colored marker. | Prepare presentation‑ready Excel charts where marker styles convey additional meaning without extra legends.
// AI Prompts: Write C# code with Aspose.Cells that sets circular markers with a red border and yellow fill for a line series in a combo chart. | Show how to change the third point of a line series to a larger square marker with a blue border and light‑green fill using Aspose.Cells. | Explain the steps to access individual ChartPoint objects and modify their marker properties in Aspose.Cells .NET.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartMarkerDemo
{
    // Creates an Excel workbook with sales and target data, adds a combo chart (column + line), applies circular red‑bordered yellow markers to the entire line series, and emphasizes a chosen point with a larger blue‑bordered green square before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                // Column A – Categories
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                // Column B – Column series values
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                // Column C – Line series values (e.g., Target)
                sheet.Cells["C1"].PutValue("Target");
                sheet.Cells["C2"].PutValue(130);
                sheet.Cells["C3"].PutValue(140);
                sheet.Cells["C4"].PutValue(170);
                sheet.Cells["C5"].PutValue(210);

                // Add a combo chart (initially a column chart)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the column series (first series)
                chart.NSeries.Add("B2:B5", true);
                // Add the line series (second series) and set its chart type to Line
                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Type = ChartType.Line; // Convert second series to line

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A5";

                // Customize marker style for the line series (second series)
                Series lineSeries = chart.NSeries[1];
                lineSeries.Marker.MarkerStyle = ChartMarkerType.Circle;   // Circular markers
                lineSeries.Marker.MarkerSize = 10;                        // Size in points
                lineSeries.Marker.ForegroundColor = Color.Red;           // Marker border color
                lineSeries.Marker.BackgroundColor = Color.Yellow;        // Marker fill color

                // Highlight specific key data points within the line series
                // Example: highlight the third point (index 2) with a different style
                if (lineSeries.Points.Count > 2)
                {
                    ChartPoint keyPoint = lineSeries.Points[2];
                    keyPoint.Marker.MarkerStyle = ChartMarkerType.Square; // Square marker for emphasis
                    keyPoint.Marker.MarkerSize = 14;                      // Larger size
                    keyPoint.Marker.ForegroundColor = Color.Blue;        // Different border color
                    keyPoint.Marker.BackgroundColor = Color.LightGreen; // Different fill color
                }

                // Save the workbook
                string outputPath = "ComboChartWithCustomMarkers.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
