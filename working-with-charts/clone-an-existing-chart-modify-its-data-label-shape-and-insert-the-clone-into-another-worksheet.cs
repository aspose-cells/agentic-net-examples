// Title: Clone a Chart, Change Its Data‑Label Shape, and Place It on Another Worksheet – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy a chart from a source worksheet, modify the cloned series’ data‑label shape, and insert the cloned chart into a destination worksheet using Aspose.Cells for .NET. The example creates a column chart, enables rectangular data labels, clones the chart via ChartShape.AddCopy, changes the label shape to an ellipse, and saves the workbook.
// Keywords: Aspose.Cells clone chart C# | ChartShape AddCopy example | modify data label shape Aspose.Cells | copy chart to another worksheet .NET | DataLabelShapeType ellipse | duplicate chart Aspose.Cells | C# Aspose.Cells chart manipulation
// Common Searches: how to duplicate a chart with Aspose.Cells for .NET | copy chart to a different worksheet using ChartShape | change data label shape of a cloned chart Aspose.Cells | Aspose.Cells C# example for ChartShape.AddCopy | clone chart and edit data labels in Excel workbook
// Developer Intent: Copy an existing chart, alter the cloned series’ data‑label shape, and embed the clone in another worksheet.
// Use Cases: Create a master chart template and reuse it across multiple sheets with customized label shapes for comparative dashboards. | Generate a summary report that aggregates source charts, each cloned and styled differently for visual emphasis. | Automate workbook production where charts are duplicated per department and their data‑label shapes are tailored to regional branding.
// AI Prompts: Show C# code that clones a chart from one worksheet to another and sets the cloned series data labels to an ellipse using Aspose.Cells. | Explain step‑by‑step how to use ChartShape.AddCopy to copy a chart and then change DataLabelShapeType without affecting the original chart. | Provide a concise Aspose.Cells for .NET example that copies a chart, modifies its data‑label shape, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartCloneDemo
{
    // Demonstrates how to copy a chart from a source worksheet, modify the cloned series’ data‑label shape, and insert the cloned chart into a destination worksheet using Aspose.Cells for .NET. The example creates a column chart, enables rectangular data labels, clones the chart via ChartShape.AddCopy, changes the label shape to an ellipse, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- Source worksheet with original chart ----------
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Populate sample data
            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["B1"].PutValue("Value");
            srcSheet.Cells["A2"].PutValue("A");
            srcSheet.Cells["A3"].PutValue("B");
            srcSheet.Cells["A4"].PutValue("C");
            srcSheet.Cells["B2"].PutValue(10);
            srcSheet.Cells["B3"].PutValue(20);
            srcSheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart srcChart = srcSheet.Charts[chartIdx];
            srcChart.NSeries.Add("B2:B4", true);
            srcChart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and set an initial shape type
            Series srcSeries = srcChart.NSeries[0];
            srcSeries.DataLabels.ShowValue = true;
            srcSeries.DataLabels.ShapeType = DataLabelShapeType.Rect;

            // ---------- Destination worksheet ----------
            Worksheet destSheet = workbook.Worksheets.Add("Destination");

            // Clone the chart by copying its ChartShape to the destination worksheet
            // ChartObject returns the ChartShape that represents the chart as a shape
            ChartShape srcChartShape = (ChartShape)srcChart.ChartObject;

            // AddCopy copies the shape (including the embedded chart) to the target worksheet
            // Parameters: source shape, top row, top offset (pixels), left column, left offset (pixels)
            Shape copiedShape = destSheet.Shapes.AddCopy(srcChartShape, 20, 0, 0, 0);

            // Cast the copied shape back to ChartShape to access the cloned Chart object
            ChartShape clonedChartShape = (ChartShape)copiedShape;
            Chart clonedChart = clonedChartShape.Chart;

            // Modify the data label shape of the cloned chart
            Series clonedSeries = clonedChart.NSeries[0];
            clonedSeries.DataLabels.ShapeType = DataLabelShapeType.Ellipse; // Change to ellipse

            // Optionally reposition the cloned chart (using Chart.Move if needed)
            // clonedChart.Move(20, 0, 30, 10);

            // Save the workbook
            workbook.Save("ChartCloneWithModifiedDataLabelShape.xlsx");
        }
    }
}
