// Title: C# – Resize an Aspose.Cells chart to 500 × 300 pts and anchor it at cell D5
// Description: Creates a new workbook, adds sample data, inserts a column chart, then sets the chart's WidthPt to 500 and HeightPt to 300. The chart is moved so its upper‑left corner aligns with cell D5, and the file is saved as an Excel workbook.
// Keywords: Aspose.Cells | C# | chart resize points | chart position cell | WidthPt | HeightPt | ChartObject.Move | cell D5 | 500 points | 300 points | column chart | Excel automation
// Common Searches: Aspose.Cells set chart width in points | How to move a chart to a specific cell in .NET | Resize chart to 500x300 points Aspose.Cells | Place chart at D5 using Aspose.Cells C# | ChartObject WidthPt HeightPt example
// Developer Intent: Resize a chart to 500 pts × 300 pts and place its top‑left corner in cell D5.
// Use Cases: Design a dashboard where charts must match exact dimensions. | Align several charts on a worksheet by anchoring each to a predefined cell. | Generate Excel reports that fit a pre‑designed PDF template with fixed chart sizes.
// AI Prompts: Provide C# code that resizes an Aspose.Cells chart to 500 × 300 points and anchors it at cell D5. | Explain how WidthPt, HeightPt, and ChartObject.Move work together to control chart size and placement in Aspose.Cells. | Show a step‑by‑step example of positioning a chart at a specific worksheet cell using zero‑based indices.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartResizeAndPosition
{
    // Creates a new workbook, adds sample data, inserts a column chart, then sets the chart's WidthPt to 500 and HeightPt to 300. The chart is moved so its upper‑left corner aligns with cell D5, and the file is saved as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Add some sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a chart (initial position is temporary; we'll resize and move it later)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 0, 0, 10, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Resize the chart: width = 500 points, height = 300 points
            // Use WidthPt and HeightPt properties (points) of the ChartShape object
            chart.ChartObject.WidthPt = 500;
            chart.ChartObject.HeightPt = 300;

            // Position the chart so that its upper‑left corner is at cell D5
            // Zero‑based indices: column D = 3, row 5 = 4
            // Move(topRow, leftColumn, bottomRow, rightColumn)
            // Bottom/right rows/columns can be the same as top/left because size is controlled by WidthPt/HeightPt
            chart.Move(4, 3, 4, 3);

            // Save the workbook
            workbook.Save("ResizedAndPositionedChart.xlsx");
        }
    }
}
