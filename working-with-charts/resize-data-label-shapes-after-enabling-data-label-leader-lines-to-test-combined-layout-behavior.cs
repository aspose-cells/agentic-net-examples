// Title: Resize Data Label Shapes with Leader Lines in an Aspose.Cells Column Chart (C#/.NET)
// Description: This example creates a workbook, adds a column chart, enables leader lines with custom styling, shows data labels outside the columns, disables auto‑fit for each point, and sets a fixed pixel width, height, and rectangular shape for the data label shapes before saving the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | column chart | data labels | leader lines | fixed label size | pixel width | pixel height | disable auto fit | rectangle data label shape | ResizeDataLabelShapesDemo
// Common Searches: Aspose.Cells set fixed size for chart data labels | Enable leader lines and customize data label shape in C# | Resize data label shapes after turning on leader lines Aspose.Cells | How to prevent auto‑fit of data labels in Aspose chart | Change data label shape to rectangle in Aspose.Cells .NET
// Developer Intent: Apply leader lines to a chart series and give every data label a constant pixel width, height, and rectangular shape while disabling automatic resizing.
// Use Cases: Generate Excel reports with column charts where data labels have uniform rectangular shapes for a clean, predictable layout. | Create dashboards that require leader lines and fixed‑size data labels to maintain alignment across multiple series regardless of label text length. | Automate workbook creation where consistent visual spacing of data labels is essential for printing or PDF export.
// AI Prompts: Show C# code using Aspose.Cells to enable leader lines on a column chart series and set a fixed pixel width and height for each data label, disabling auto‑fit. | Provide an Aspose.Cells example that changes all data label shapes to rectangles and customizes leader line style in .NET. | Explain how to position data labels outside data points, keep their size constant, and ensure they stay aligned when the chart is resized.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a column chart, enables leader lines with custom styling, shows data labels outside the columns, disables auto‑fit for each point, and sets a fixed pixel width, height, and rectangular shape for the data label shapes before saving the file as an Excel workbook.
class ResizeDataLabelShapesDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Set values
        chart.NSeries.CategoryData = "A2:A4";      // Set categories

        // Access the first series in the chart
        Series series = chart.NSeries[0];

        // Enable leader lines for the series and customize their appearance
        series.HasLeaderLines = true;
        series.LeaderLines.IsAuto = false;
        series.LeaderLines.Style = LineType.Solid;
        series.LeaderLines.WeightPt = 1.0;
        series.LeaderLines.Color = Color.DarkGray;

        // Show data labels and place them outside the data points
        series.DataLabels.ShowValue = true;
        series.DataLabels.Position = LabelPositionType.OutsideEnd;

        // For each data point, disable auto‑fit and set a custom size for the label shape
        foreach (ChartPoint point in series.Points)
        {
            point.DataLabels.IsResizeShapeToFitText = false; // Prevent auto‑resizing
            point.DataLabels.WidthPixel = 60;                // Custom width (pixels)
            point.DataLabels.HeightPixel = 30;               // Custom height (pixels)
            point.DataLabels.ShapeType = DataLabelShapeType.Rect; // Use a rectangle shape
        }

        // Save the workbook with the configured chart
        workbook.Save("ResizeDataLabelShapesDemo.xlsx");
    }
}
