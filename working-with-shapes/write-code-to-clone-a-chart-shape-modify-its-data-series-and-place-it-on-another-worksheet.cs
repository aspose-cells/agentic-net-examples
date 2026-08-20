// Title: Clone and modify a chart shape across worksheets with Aspose.Cells for .NET
// Description: Creates a source worksheet with sample data, adds a column chart, copies its ChartShape to a destination sheet using Shapes.AddCopy, casts the copy back to ChartShape, renames the original series, adds new data and a new series on the target sheet, and saves the workbook with the updated chart.
// Keywords: Aspose.Cells | .NET | C# | clone chart shape | copy chart worksheet | Shapes.AddCopy | ChartShape | modify NSeries | add chart series | Excel chart automation
// Common Searches: Aspose.Cells copy chart to another sheet | How to duplicate a chart shape in C# | Change series of a cloned chart Aspose.Cells | Shapes.AddCopy example for charts | Clone chart and edit data range .NET
// Developer Intent: Duplicate an existing chart, adjust its series, and place the modified chart on a different worksheet.
// Use Cases: Reuse a master chart template on multiple report sheets with customized data series. | Automate dashboard creation by cloning a chart and appending new series per worksheet. | Generate region‑specific charts by copying a base chart and adjusting its NSeries for each locale.
// AI Prompts: Write C# code using Aspose.Cells to copy a ChartShape from one worksheet to another and rename its series. | Show how to add a new NSeries to a cloned chart on a different worksheet with Aspose.Cells for .NET. | Explain the use of Shapes.AddCopy for duplicating a chart and then modifying its data ranges.

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a source worksheet with sample data, adds a column chart, copies its ChartShape to a destination sheet using Shapes.AddCopy, casts the copy back to ChartShape, renames the original series, adds new data and a new series on the target sheet, and saves the workbook with the updated chart.
class CloneChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet srcSheet = workbook.Worksheets[0];
        srcSheet.Name = "Source";

        // Populate sample data for the source chart
        srcSheet.Cells["A1"].PutValue("Category");
        srcSheet.Cells["B1"].PutValue("Value");
        srcSheet.Cells["A2"].PutValue("A");
        srcSheet.Cells["B2"].PutValue(10);
        srcSheet.Cells["A3"].PutValue("B");
        srcSheet.Cells["B3"].PutValue(20);
        srcSheet.Cells["A4"].PutValue("C");
        srcSheet.Cells["B4"].PutValue(30);

        // Add a chart to the source worksheet
        int chartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart srcChart = srcSheet.Charts[chartIdx];
        srcChart.NSeries.Add("B2:B4", true);
        srcChart.NSeries.CategoryData = "A2:A4";

        // Add a destination worksheet where the cloned chart will be placed
        Worksheet destSheet = workbook.Worksheets.Add("Destination");

        // Obtain the ChartShape (visual representation) of the source chart
        ChartShape srcChartShape = srcChart.ChartObject;

        // Copy the chart shape to the destination worksheet.
        // Parameters: topRow, top (pixel offset), leftColumn, left (pixel offset)
        Shape copiedShape = destSheet.Shapes.AddCopy(srcChartShape, 5, 0, 0, 0);

        // Cast the copied shape back to ChartShape to access its Chart object
        ChartShape destChartShape = (ChartShape)copiedShape;
        Chart destChart = destChartShape.Chart;

        // Modify the data series of the cloned chart
        // Example: rename the existing series
        if (destChart.NSeries.Count > 0)
        {
            destChart.NSeries[0].Name = "Original Series Modified";
        }

        // Add new data on the destination sheet for an additional series
        destSheet.Cells["A6"].PutValue("D");
        destSheet.Cells["B6"].PutValue(40);
        destSheet.Cells["A7"].PutValue("E");
        destSheet.Cells["B7"].PutValue(50);

        // Add a new series that references the new data range
        destChart.NSeries.Add("B6:B7", true);
        destChart.NSeries[destChart.NSeries.Count - 1].Name = "New Series";

        // Save the workbook with the cloned and modified chart
        workbook.Save("ClonedChart.xlsx");
    }
}
