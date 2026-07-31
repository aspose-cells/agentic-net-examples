// Title: Fix Data Label Shape Size After Enabling Leader Lines in an Aspose.Cells Column Chart (C#)
// Description: Creates a workbook, adds a column chart, turns on leader lines for the first series, applies a right‑arrow callout data label, disables automatic shape resizing, and sets the label shape to 60 px × 30 px before saving the file.
// Keywords: Aspose.Cells | C# chart | data label shape size | leader lines | right arrow callout | fixed width height | ResizeDataLabelShapesDemo | column chart | customize leader line | IsResizeShapeToFitText
// Common Searches: Aspose.Cells set fixed width height for data label | enable leader lines and custom data label size C# | disable automatic data label resizing Aspose.Cells | right arrow callout data label chart Aspose | ResizeDataLabelShapesDemo example
// Developer Intent: Set a fixed width and height for data label shapes while leader lines are active.
// Use Cases: Generate a column chart with leader lines and a right‑arrow callout whose label shape stays at 60 × 30 px. | Maintain consistent label layout across multiple series by turning off auto‑fit resizing. | Apply custom leader line style, weight, and color together with fixed data label dimensions for reporting templates. | Create reusable chart templates where data label shapes have predetermined size regardless of content.
// AI Prompts: Write C# code using Aspose.Cells to add a column chart, enable leader lines, and set data label shape to a specific width and height. | Show how to disable automatic resizing of data label shapes and assign explicit dimensions after turning on leader lines in an Aspose.Cells chart. | Provide an example that customizes leader line style and keeps data label callout shapes at a fixed size. | Explain the effect of IsResizeShapeToFitText false on data label layout when using callout shapes.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// Creates a workbook, adds a column chart, turns on leader lines for the first series, applies a right‑arrow callout data label, disables automatic shape resizing, and sets the label shape to 60 px × 30 px before saving the file.
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the first series
        Series series = chart.NSeries[0];

        // Enable leader lines for the series
        series.HasLeaderLines = true;
        // Optional: customize leader line appearance
        series.LeaderLines.IsAuto = false;
        series.LeaderLines.Style = LineType.Solid;
        series.LeaderLines.WeightPt = 1.0;
        series.LeaderLines.Color = Color.DarkGray;

        // Enable data labels and set a callout shape type
        series.DataLabels.ShowValue = true;
        series.DataLabels.ShapeType = DataLabelShapeType.RightArrowCallout;

        // Disable automatic resizing of the label shape to fit text
        series.DataLabels.IsResizeShapeToFitText = false;
        // Set explicit size for the data label shape (smaller than auto‑fit size)
        series.DataLabels.Width = 60;   // width in pixels
        series.DataLabels.Height = 30;  // height in pixels

        // Save the workbook
        workbook.Save("ResizeDataLabelShapesDemo.xlsx");
    }
}
