// Title: Clone a column chart, change its data label shape to ellipse, and place the clone on another worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to copy a ChartShape from a source worksheet to a destination worksheet, then set the cloned series DataLabels.ShapeType to Ellipse. | Programmatically duplicate an existing chart on a different sheet with ShapeCollection.AddCopy and modify the cloned chart’s data label shape via DataLabelShapeType.
// Common Searches: Aspose.Cells C# copy chart to another worksheet and change data label shape | How to duplicate a chart and set data label shape to ellipse in Aspose.Cells | Clone chart object between worksheets using ShapeCollection.AddCopy Aspose.Cells .NET | Change data label shape of a cloned chart in Aspose.Cells C# example
// Tags: ChartShape AddCopy Aspose.Cells C# | clone chart between worksheets Aspose.Cells | ellipse DataLabelShapeType Aspose.Cells | modify cloned chart series data labels Aspose.Cells | column chart cloning Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing; // for DataLabelShapeType

namespace AsposeCellsChartCloneDemo
{
    // The example creates a workbook with sample data, adds a column chart on a source sheet with rectangular data labels, clones the chart to a destination sheet using ShapeCollection.AddCopy, changes the cloned series' data label shape to an ellipse, and saves the file as ChartCloneWithModifiedDataLabelShape.xlsx.
    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook and populate data ----------
            Workbook workbook = new Workbook();
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Sample data
            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["B1"].PutValue("Value");
            srcSheet.Cells["A2"].PutValue("A");
            srcSheet.Cells["A3"].PutValue("B");
            srcSheet.Cells["A4"].PutValue("C");
            srcSheet.Cells["B2"].PutValue(10);
            srcSheet.Cells["B3"].PutValue(20);
            srcSheet.Cells["B4"].PutValue(30);

            // ---------- Add a chart to the source sheet ----------
            int chartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart srcChart = srcSheet.Charts[chartIdx];
            srcChart.NSeries.Add("B2:B4", true);
            srcChart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series (so we can see shape change later)
            Series srcSeries = srcChart.NSeries[0];
            srcSeries.DataLabels.ShowValue = true;
            srcSeries.DataLabels.ShapeType = DataLabelShapeType.Rect; // initial shape

            // ---------- Add a destination worksheet ----------
            Worksheet destSheet = workbook.Worksheets.Add("Destination");

            // ---------- Clone the chart by copying its ChartShape ----------
            // ChartObject returns the ChartShape that represents the chart on the sheet
            ChartShape srcChartShape = (ChartShape)srcChart.ChartObject;

            // Use ShapeCollection.AddCopy to copy the chart shape to the destination sheet
            // Position the cloned chart at rows 5-15 and columns 2-7 (adjust as needed)
            Shape copiedShape = destSheet.Shapes.AddCopy(srcChartShape, 5, 0, 2, 0);

            // Cast the copied shape back to ChartShape to access the underlying Chart
            ChartShape clonedChartShape = (ChartShape)copiedShape;
            Chart clonedChart = clonedChartShape.Chart;

            // ---------- Modify data label shape of the cloned chart ----------
            // Ensure the series exists (it will be a copy of the original series)
            Series clonedSeries = clonedChart.NSeries[0];
            clonedSeries.DataLabels.ShowValue = true;
            // Change the shape type, e.g., to Ellipse
            clonedSeries.DataLabels.ShapeType = DataLabelShapeType.Ellipse;

            // ---------- Save the workbook ----------
            workbook.Save("ChartCloneWithModifiedDataLabelShape.xlsx");
        }
    }
}
