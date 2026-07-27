// Title: Clone and Modify a ChartShape in Aspose.Cells for .NET – Copy to Another Worksheet, Add Series, Change Type
// Description: C# example that creates a workbook, adds a column chart, clones the ChartShape onto a new worksheet, appends an extra data series from column C, reorders the series, switches the chart to a line type, and saves the file as ClonedChartExample.xlsx.
// Keywords: Aspose.Cells C# chart clone | ChartShape AddCopy example | copy chart between worksheets .NET | modify chart NSeries programmatically | add data series to cloned chart | change chart type Aspose.Cells | Aspose.Cells chart manipulation | ChartShape cloning tutorial | Aspose.Cells API chart series | C# Excel chart automation
// Common Searches: how to clone a chart shape with Aspose.Cells | copy chart to another worksheet C# | add new series to a cloned chart Aspose.Cells | change chart type after copying chart shape | move series order in Aspose.Cells chart | Aspose.Cells ChartShape AddCopy usage
// Developer Intent: Copy a chart shape to a different worksheet, adjust its series, and change its chart type using Aspose.Cells for .NET.
// Use Cases: Create a template chart once and reuse it on multiple sheets with customized data series. | Generate comparative visualizations by cloning an existing chart and adding supplemental data. | Automate report generation where the same chart layout is needed across several worksheets with different data sources.
// AI Prompts: Generate C# code that uses Aspose.Cells to copy a ChartShape from one worksheet to another, then adds a new series from column C and sets the chart type to Line. | Explain how to reorder data series in a cloned chart's NSeries collection with Aspose.Cells for .NET. | Show the steps to save an Excel file after cloning and modifying a chart shape using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartCloneExample
{
    // C# example that creates a workbook, adds a column chart, clones the ChartShape onto a new worksheet, appends an extra data series from column C, reorders the series, switches the chart to a line type, and saves the file as ClonedChartExample.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Populate sample data for the original chart
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["A4"].PutValue("C");
            sourceSheet.Cells["B1"].PutValue("Value");
            sourceSheet.Cells["B2"].PutValue(10);
            sourceSheet.Cells["B3"].PutValue(20);
            sourceSheet.Cells["B4"].PutValue(30);
            sourceSheet.Cells["C1"].PutValue("Extra");
            sourceSheet.Cells["C2"].PutValue(15);
            sourceSheet.Cells["C3"].PutValue(25);
            sourceSheet.Cells["C4"].PutValue(35);

            // Add a column chart to the source worksheet
            int chartIndex = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart sourceChart = sourceSheet.Charts[chartIndex];
            sourceChart.NSeries.Add("B2:B4", true);               // Values
            sourceChart.NSeries.CategoryData = "A2:A4";           // Categories
            sourceChart.NSeries[0].Name = "Original Series";

            // Obtain the ChartShape (the visual shape of the chart)
            ChartShape sourceChartShape = sourceChart.ChartObject;

            // Add a new worksheet where the cloned chart will be placed
            Worksheet destSheet = workbook.Worksheets.Add("ClonedChartSheet");

            // Clone the chart shape onto the destination worksheet
            // Parameters: source shape, top row, top offset (pixels), left column, left offset (pixels)
            Shape clonedShape = destSheet.Shapes.AddCopy(sourceChartShape, 5, 0, 0, 0);

            // Cast the cloned shape back to ChartShape to access its Chart object
            ChartShape clonedChartShape = (ChartShape)clonedShape;
            Chart clonedChart = clonedChartShape.Chart;

            // ---- Modify the cloned chart's data series ----

            // Add an additional series to the cloned chart using data from column C
            clonedChart.NSeries.Add("C2:C4", false);
            clonedChart.NSeries[1].Name = "Added Series";

            // Move the newly added series to the first position (so it appears before the original series)
            clonedChart.NSeries[1].Move(-1); // Moves up by one position

            // Optionally change the chart type of the cloned chart
            clonedChart.Type = ChartType.Line;

            // Save the workbook to a file
            workbook.Save("ClonedChartExample.xlsx");
        }
    }
}
