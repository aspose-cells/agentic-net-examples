// Title: Aspose.Cells for .NET: Create a Waterfall Chart with Separate Colors for Start, Intermediate, and Total Points (C#)
// Description: This C# example builds a workbook, adds category and value data, inserts a Waterfall chart, links the series to the range, and colors the first bar green (start), the middle bars blue (intermediate), and the final bar red (total) before saving as WaterfallChartDemo.xlsx.
// Keywords: Aspose.Cells | C# | Waterfall chart | custom point colors | start bar green | intermediate bars blue | total bar red | .NET chart formatting | chart point coloring | Aspose.Cells example
// Common Searches: Aspose.Cells set color for first point in waterfall chart | C# change individual bar colors in Aspose.Cells waterfall | How to highlight start and total values in Aspose.Cells chart | Waterfall chart custom colors Aspose.Cells .NET | Apply different colors to waterfall series points using Aspose.Cells
// Developer Intent: The developer wants to generate a waterfall chart and programmatically assign distinct colors to the start bar, each intermediate bar, and the total bar.
// Use Cases: Financial statements where the opening balance appears in green, period adjustments in blue, and the closing balance in red for quick visual comparison. | Project budget waterfall that distinguishes the initial allocation, incremental changes, and final total with custom bar colors. | Sales performance waterfall highlighting the baseline sales figure in green, quarterly variations in blue, and the final target in red.
// AI Prompts: Generate C# code with Aspose.Cells that creates a waterfall chart and colors the start point green, intermediate points blue, and total point red. | Explain how to programmatically set individual point colors in an Aspose.Cells waterfall chart based on their index. | Provide a step‑by‑step tutorial for applying separate colors to start, intermediate, and total bars in a waterfall chart using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsWaterfallDemo
{
    // This C# example builds a workbook, adds category and value data, inserts a Waterfall chart, links the series to the range, and colors the first bar green (start), the middle bars blue (intermediate), and the final bar red (total) before saving as WaterfallChartDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the waterfall chart
            // Column A: Categories (Start, Intermediate points, Total)
            // Column B: Values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(5000);   // Starting value

            sheet.Cells["A3"].PutValue("Q1");
            sheet.Cells["B3"].PutValue(2000);   // Intermediate increase

            sheet.Cells["A4"].PutValue("Q2");
            sheet.Cells["B4"].PutValue(-1500);  // Intermediate decrease

            sheet.Cells["A5"].PutValue("Q3");
            sheet.Cells["B5"].PutValue(3000);   // Intermediate increase

            sheet.Cells["A6"].PutValue("Total");
            sheet.Cells["B6"].PutValue(8500);   // Total value

            // Add a Waterfall chart
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 8, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Apply distinct colors:
            // First point (Start) - Green
            ChartPoint startPoint = chart.NSeries[0].Points[0];
            startPoint.Area.ForegroundColor = Color.Green;

            // Intermediate points (indexes 1 to 4) - Blue
            for (int i = 1; i <= 4; i++)
            {
                ChartPoint intermediatePoint = chart.NSeries[0].Points[i];
                intermediatePoint.Area.ForegroundColor = Color.Blue;
            }

            // Last point (Total) - Red
            ChartPoint totalPoint = chart.NSeries[0].Points[5];
            totalPoint.Area.ForegroundColor = Color.Red;

            // Save the workbook
            workbook.Save("WaterfallChartDemo.xlsx");
        }
    }
}
