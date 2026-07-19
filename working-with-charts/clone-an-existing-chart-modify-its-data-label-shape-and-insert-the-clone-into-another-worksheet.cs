// Title: Clone a chart, modify its data‑label shape, and place it on another worksheet using Aspose.Cells for .NET
// Description: Shows how to build a workbook, add a column chart with rectangular data labels, copy the chart to a new worksheet via ChartShape.AddCopy, change the cloned series’ data‑label shape to an ellipse, and save the file as an Excel workbook.
// Keywords: Aspose.Cells | C# chart cloning | ChartShape AddCopy | data label shape | DataLabelShapeType ellipse | copy chart between worksheets | Aspose.Cells .NET example | Excel chart duplication
// Common Searches: how to duplicate a chart with Aspose.Cells | copy chart to another worksheet C# Aspose.Cells | change data label shape of a cloned chart | ChartShape AddCopy parameters example | Aspose.Cells modify data labels ellipse
// Developer Intent: The developer needs to replicate an existing chart, adjust the shape of its data‑label markers, and insert the replicated chart into a different worksheet programmatically.
// Use Cases: Create a summary sheet that reuses a source chart but displays ellipse‑shaped data labels for a cleaner look. | Generate multi‑sheet reports where each sheet contains a copy of the same chart with distinct label styling to differentiate sections. | Automate the migration of charts across worksheets while applying corporate branding to data‑label shapes.
// AI Prompts: Write C# code with Aspose.Cells that clones a chart from one worksheet to another and sets the cloned series’ data labels to an ellipse shape. | Explain the parameters of ChartShape.AddCopy and how they affect the position of a copied chart on the target worksheet. | Provide a loop that iterates through all series in a cloned chart and applies a custom DataLabelShapeType to each series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartCloneDemo
{
    // Shows how to build a workbook, add a column chart with rectangular data labels, copy the chart to a new worksheet via ChartShape.AddCopy, change the cloned series’ data‑label shape to an ellipse, and save the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Prepare data in the first worksheet (source chart)
            // -------------------------------------------------
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Value");
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["A4"].PutValue("C");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["B4"].PutValue(30);

            // Add a chart to the source worksheet
            int chartIdx = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart sourceChart = sourceSheet.Charts[chartIdx];
            sourceChart.NSeries.Add("B2:B4", true);
            sourceChart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and set an initial shape type
            Series srcSeries = sourceChart.NSeries[0];
            srcSeries.DataLabels.ShowValue = true;
            srcSeries.DataLabels.ShapeType = DataLabelShapeType.Rect;

            // -------------------------------------------------
            // Add a second worksheet where the cloned chart will be placed
            // -------------------------------------------------
            Worksheet targetSheet = workbook.Worksheets.Add("ClonedChartSheet");

            // -------------------------------------------------
            // Clone the chart by copying its ChartShape to the target sheet
            // -------------------------------------------------
            // Obtain the ChartShape that represents the source chart
            ChartShape sourceChartShape = (ChartShape)sourceChart.ChartObject;

            // Use ShapeCollection.AddCopy to copy the chart shape to the target sheet
            // Parameters: source shape, top row, top offset (pixels), left column, left offset (pixels)
            Shape copiedShape = targetSheet.Shapes.AddCopy(sourceChartShape, 2, 0, 2, 0);

            // Cast the copied shape back to ChartShape to access the underlying Chart object
            ChartShape clonedChartShape = (ChartShape)copiedShape;
            Chart clonedChart = clonedChartShape.Chart;

            // -------------------------------------------------
            // Modify the data label shape of the cloned chart
            // -------------------------------------------------
            // Ensure the series exists
            if (clonedChart.NSeries.Count > 0)
            {
                Series clonedSeries = clonedChart.NSeries[0];
                clonedSeries.DataLabels.ShowValue = true;                     // Show values
                clonedSeries.DataLabels.ShapeType = DataLabelShapeType.Ellipse; // Change shape
            }

            // -------------------------------------------------
            // Save the workbook (save rule)
            // -------------------------------------------------
            workbook.Save("ClonedChartWithModifiedDataLabels.xlsx");
        }
    }
}
