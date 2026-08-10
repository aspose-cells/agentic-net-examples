// Title: Aspose.Cells .NET – Radar Chart with Row‑Based (Transposed) Data, Fixed‑Size Data Labels & Axis Labels
// Description: Creates a workbook, fills category and horizontally‑arranged series data, adds a radar chart, uses a transposed (plot‑by‑row) range for the series, sets category axis labels, enables radar axis labels, shows numeric values in data labels with a fixed width and rectangular shape, disables auto‑fit, recalculates the chart, and saves the file.
// Keywords: Aspose.Cells | C# radar chart | transposed range plot by row | fixed size data labels | radar axis labels | HasRadarAxisLabels | DataLabelShapeType.Rect | NSeries.Add false vertical | chart.Calculate | .NET charting
// Common Searches: Aspose.Cells add data labels to radar chart .NET | plot radar chart by row using Aspose.Cells | set fixed width for chart data labels Aspose.Cells | enable radar axis labels in Aspose.Cells chart | disable auto‑resize of data label shapes Aspose.Cells
// Developer Intent: Add a radar chart that reads series data by row, displays value labels with a fixed rectangular shape, and shows axis labels.
// Use Cases: Display numeric values next to each point when series data is stored horizontally. | Create a radar chart where each row of a range represents a separate series (plot‑by‑row). | Maintain consistent label dimensions by turning off auto‑fit and assigning a fixed width and shape.
// AI Prompts: Generate C# code using Aspose.Cells to build a radar chart that reads series data from a transposed row range and shows fixed‑width rectangular data labels. | Show how to enable radar axis labels and disable auto‑resize for data label shapes in an Aspose.Cells radar chart. | Provide an Aspose.Cells .NET example that adds a radar chart, sets category data, uses NSeries.Add with plot‑by‑row, and configures data labels to show values with a constant size.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsRadarChartDemo
{
    // Creates a workbook, fills category and horizontally‑arranged series data, adds a radar chart, uses a transposed (plot‑by‑row) range for the series, sets category axis labels, enables radar axis labels, shows numeric values in data labels with a fixed width and rectangular shape, disables auto‑fit, recalculates the chart, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Category axis (A2:A4)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Cat1");
            sheet.Cells["A3"].PutValue("Cat2");
            sheet.Cells["A4"].PutValue("Cat3");

            // Series values placed horizontally (B2:D2, B3:D3, B4:D4)
            // This layout will be read as a transposed range (by row)
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["D1"].PutValue("Series3");

            sheet.Cells["B2"].PutValue(4);
            sheet.Cells["C2"].PutValue(5);
            sheet.Cells["D2"].PutValue(6);

            sheet.Cells["B3"].PutValue(2);
            sheet.Cells["C3"].PutValue(3);
            sheet.Cells["D3"].PutValue(4);

            sheet.Cells["B4"].PutValue(5);
            sheet.Cells["C4"].PutValue(7);
            sheet.Cells["D4"].PutValue(9);

            // Add a radar chart
            int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Use a transposed range for the series data (plot by row)
            // The range B2:D4 contains the values; isVertical = false treats rows as series
            chart.NSeries.Add("B2:D4", false);
            // Set category (axis) labels
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series (the radar chart has only one series in this example)
            Series series = chart.NSeries[0];

            // Enable radar axis labels (category labels on the radar)
            series.HasRadarAxisLabels = true;

            // Enable data labels and configure them
            series.DataLabels.ShowValue = true;                 // Show the numeric values
            series.DataLabels.IsResizeShapeToFitText = false;   // Prevent auto‑fit so we can set a fixed size
            series.DataLabels.Width = 50;                       // Example fixed width
            series.DataLabels.ShapeType = DataLabelShapeType.Rect; // Optional: set shape type

            // Recalculate the chart to apply changes (lifecycle rule)
            chart.Calculate();

            // Save the workbook (lifecycle rule)
            workbook.Save("RadarChartWithDataLabels.xlsx");
        }
    }
}
