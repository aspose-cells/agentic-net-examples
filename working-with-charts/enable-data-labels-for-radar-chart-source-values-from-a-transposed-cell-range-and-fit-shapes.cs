// Title: Aspose.Cells .NET – Add data labels to a transposed radar chart and auto‑fit label shapes
// Description: Creates a workbook, fills a table, adds a radar chart, sets the data range by rows to transpose the source, shows numeric data labels, auto‑sizes label shapes, enables radar axis labels, recalculates layout, and saves the file.
// Keywords: Aspose.Cells radar chart | C# data labels | transpose chart data range | auto‑fit label shape | radar axis labels | .NET chart programming | Excel radar chart automation
// Common Searches: Aspose.Cells show values on radar chart series | transpose chart data range by rows Aspose.Cells | auto size data label shape in radar chart .NET | add category axis labels to radar chart Aspose.Cells | C# code for radar chart with data labels
// Developer Intent: Create a radar chart from transposed data, display value labels, and automatically resize label shapes.
// Use Cases: Convert a vertical data table into a radar chart without manual transposition. | Display each data point’s value directly on the radar chart for quick analysis. | Ensure label boxes adjust to their content, keeping the chart clean and readable. | Add axis (category) labels to identify each radar dimension.
// AI Prompts: Generate C# code using Aspose.Cells to build a radar chart, transpose the source range, and enable data labels with values. | Show how to configure data label shapes to auto‑fit their text in an Aspose.Cells radar chart. | Provide an example that adds radar axis (category) labels and recalculates the chart layout with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsRadarChartDemo
{
    // Creates a workbook, fills a table, adds a radar chart, sets the data range by rows to transpose the source, shows numeric data labels, auto‑sizes label shapes, enables radar axis labels, recalculates layout, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the radar chart
            // Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Cat1");
            sheet.Cells["A3"].PutValue("Cat2");
            sheet.Cells["A4"].PutValue("Cat3");

            // Series values (will be transposed when setting the chart data range)
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(4);
            sheet.Cells["B3"].PutValue(2);
            sheet.Cells["B4"].PutValue(5);

            // Add a radar chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart data range and specify that the data should be plotted by row (transposed)
            // isVertical = false means series are plotted by row, effectively transposing the range
            chart.SetChartDataRange("A1:B4", false);

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;               // Show the numeric values
            series.DataLabels.IsResizeShapeToFitText = true;  // Auto‑fit the label shape to its text

            // Enable radar axis (category) labels
            series.HasRadarAxisLabels = true;

            // Recalculate the chart layout (optional but ensures proper positioning)
            chart.Calculate();

            // Save the workbook
            workbook.Save("RadarChartWithDataLabels.xlsx");
        }
    }
}
