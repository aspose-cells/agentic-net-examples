// Title: Clone a Chart, Change Its Data Label Shape, and Insert into Another Worksheet – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data, build a column chart with round‑rect data labels, clone the chart to a new worksheet using ChartShape.AddCopy, change the cloned series' data label shape to an ellipse, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart clone | copy chart to another worksheet | modify data label shape Aspose.Cells | ChartShape AddCopy C# | Aspose.Cells .NET chart example | duplicate chart programmatically | change data label shape ellipse | C# Aspose.Cells chart manipulation
// Common Searches: Aspose.Cells copy chart to different sheet | How to clone a chart in Aspose.Cells C# | Change data label shape of a chart with Aspose.Cells | ChartShape AddCopy usage example | C# Aspose.Cells duplicate chart and edit labels
// Developer Intent: Programmatically duplicate an existing chart, alter its data label shape, and place the copy on another worksheet using Aspose.Cells for .NET.
// Use Cases: Create a summary sheet that mirrors a source chart while applying a distinct label style. | Generate multiple versions of the same chart across worksheets for varied presentation needs. | Automate workbook generation where each sheet requires a chart with customized data label shapes.
// AI Prompts: Show C# code to clone a chart from one worksheet to another and set the cloned chart's data labels to an ellipse using Aspose.Cells. | Explain how to use ChartShape.AddCopy to duplicate a chart and modify its series properties in Aspose.Cells for .NET. | Provide step‑by‑step instructions for copying a chart, retrieving the cloned Chart object, and changing the data label shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartCloneDemo
{
    // Demonstrates how to create a workbook, add sample data, build a column chart with round‑rect data labels, clone the chart to a new worksheet using ChartShape.AddCopy, change the cloned series' data label shape to an ellipse, and save the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // First worksheet – original chart source
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["B1"].PutValue("Value");
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["A4"].PutValue("C");
            sourceSheet.Cells["B4"].PutValue(30);

            // Add a column chart to the source sheet
            int chartIndex = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart originalChart = sourceSheet.Charts[chartIndex];
            originalChart.NSeries.Add("B2:B4", true);
            originalChart.NSeries.CategoryData = "A2:A4";

            // Enable data labels and set their shape type
            Series series = originalChart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.ShapeType = DataLabelShapeType.RoundRect; // original shape

            // Obtain the ChartShape that represents the chart object
            ChartShape originalChartShape = (ChartShape)originalChart.ChartObject;

            // Add a second worksheet where the cloned chart will be placed
            Worksheet targetSheet = workbook.Worksheets.Add("ClonedChartSheet");

            // Clone the chart by copying its ChartShape to the target worksheet
            // Parameters: source shape, top row, top offset (pixels), left column, left offset (pixels)
            Shape clonedShape = targetSheet.Shapes.AddCopy(originalChartShape, 5, 0, 15, 5);

            // Retrieve the Chart object from the cloned shape
            Chart clonedChart = ((ChartShape)clonedShape).Chart;

            // Modify the data label shape of the cloned chart
            clonedChart.NSeries[0].DataLabels.ShapeType = DataLabelShapeType.Ellipse;

            // Save the workbook (save rule)
            workbook.Save("ClonedChartDemo.xlsx");
        }
    }
}
