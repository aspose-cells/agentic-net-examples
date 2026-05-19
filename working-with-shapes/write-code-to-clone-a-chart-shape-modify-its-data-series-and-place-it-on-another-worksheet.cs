using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Populate source sheet with sample data
            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["B1"].PutValue("Value");
            srcSheet.Cells["A2"].PutValue("A");
            srcSheet.Cells["A3"].PutValue("B");
            srcSheet.Cells["A4"].PutValue("C");
            srcSheet.Cells["B2"].PutValue(10);
            srcSheet.Cells["B3"].PutValue(20);
            srcSheet.Cells["B4"].PutValue(30);

            // Add a chart to the source sheet
            int chartIndex = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart srcChart = srcSheet.Charts[chartIndex];
            srcChart.NSeries.Add("B2:B4", true);          // Values
            srcChart.NSeries.CategoryData = "A2:A4";     // Categories

            // Access the ChartShape (the shape that represents the chart)
            ChartShape srcChartShape = srcChart.ChartObject;

            // Add a second worksheet where the cloned chart will be placed
            Worksheet destSheet = workbook.Worksheets.Add("Destination");

            // Populate destination sheet with its own data (will be used for the cloned chart)
            destSheet.Cells["A1"].PutValue("Category");
            destSheet.Cells["B1"].PutValue("Value");
            destSheet.Cells["A2"].PutValue("X");
            destSheet.Cells["A3"].PutValue("Y");
            destSheet.Cells["A4"].PutValue("Z");
            destSheet.Cells["B2"].PutValue(40);
            destSheet.Cells["B3"].PutValue(50);
            destSheet.Cells["B4"].PutValue(60);

            // Clone the chart shape onto the destination worksheet.
            // Parameters: source shape, top row, top offset (pixels), left column, left offset (pixels)
            Shape clonedShape = destSheet.Shapes.AddCopy(srcChartShape, 5, 0, 0, 0);

            // Cast the cloned shape back to ChartShape to access the embedded chart
            ChartShape clonedChartShape = (ChartShape)clonedShape;
            Chart clonedChart = clonedChartShape.Chart;

            // Modify the cloned chart's data series to point to the data on the destination sheet
            // Here we replace the existing series with a new one that uses the destination data range
            clonedChart.NSeries.Clear(); // Remove the series copied from the source chart
            clonedChart.NSeries.Add("B2:B4", true);          // New values from destination sheet
            clonedChart.NSeries[0].Name = "Cloned Series"; // Optional: set series name
            clonedChart.NSeries.CategoryData = "A2:A4";    // New categories from destination sheet

            // Optionally, move the cloned chart to a different position on the destination sheet
            clonedChart.Move(10, 2, 20, 8); // topRow, leftColumn, bottomRow, rightColumn

            // Save the workbook
            workbook.Save("ClonedChartExample.xlsx");
        }
    }
}