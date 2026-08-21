// Title: C# – Resize Aspose.Cells Chart to 500 pt × 300 pt and Position at Cell D5
// Description: Creates a new workbook, adds sample data, inserts a column chart, sets its data range, resizes the chart to 500 points wide and 300 points high, and moves the chart so its upper‑left corner aligns with cell D5 before saving as an XLSX file.
// Keywords: Aspose.Cells | C# chart resize | chart width points | chart height points | move chart to cell D5 | Excel chart positioning | .NET workbook chart | ResizeChart.xlsx
// Common Searches: Aspose.Cells set chart size in points | How to move a chart to a specific cell in .NET | Resize and reposition Excel chart programmatically | C# Aspose.Cells chart dimensions | Place chart at cell D5 using Aspose
// Developer Intent: Resize a chart to 500 pt × 300 pt and anchor its upper‑left corner to cell D5 in a .NET workbook.
// Use Cases: Generate a column chart from worksheet data and enforce exact dimensions for consistent report layouts. | Align charts with specific cells (e.g., D5) to match template designs. | Dynamically adjust chart size based on data volume while preserving placement.
// AI Prompts: Write C# code with Aspose.Cells that resizes a chart to 500 points wide, 300 points high, and moves it to cell D5. | Explain the difference between ChartObject.WidthPt/HeightPt and the Move method for chart positioning in Aspose.Cells. | Show an alternative way to position a chart using UpperLeftRow and UpperLeftColumn properties.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartResize
{
    // Creates a new workbook, adds sample data, inserts a column chart, sets its data range, resizes the chart to 500 points wide and 300 points high, and moves the chart so its upper‑left corner aligns with cell D5 before saving as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data for the chart (required for a valid chart)
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart (initial position is arbitrary)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Resize the chart: width = 500 points, height = 300 points
            chart.ChartObject.WidthPt = 500;   // Width in points
            chart.ChartObject.HeightPt = 300;  // Height in points

            // Move the chart so that its upper‑left corner aligns with cell D5
            // D5 corresponds to row index 4 (zero‑based) and column index 3
            int bottomRow = chart.ChartObject.LowerRightRow;
            int rightColumn = chart.ChartObject.LowerRightColumn;
            chart.Move(4, 3, bottomRow, rightColumn);

            // Save the workbook with the resized and repositioned chart
            workbook.Save("ResizedChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
