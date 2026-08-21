// Title: C# – Read and Toggle Chart Legend Position (Opposite Corner) with Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart, reads the legend's current Position, switches it to the opposite side (left↔right, top↔bottom, otherwise Corner), and saves the file as LegendPositionSwitched.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells chart legend position C# | toggle legend position Aspose.Cells | move legend opposite corner .NET | LegendPositionType example | C# Excel chart legend manipulation | Aspose.Cells change legend side | programmatic legend placement
// Common Searches: Aspose.Cells get current legend position C# | C# switch chart legend from left to right Aspose.Cells | how to move Excel chart legend to opposite corner using Aspose.Cells | toggle chart legend top bottom Aspose.Cells .NET | Aspose.Cells legend position example
// Developer Intent: Read a chart's legend position, set it to the opposite corner, and persist the change in the workbook programmatically.
// Use Cases: Dynamic report generation where legend placement adapts to chart size or layout. | Template-driven Excel exports that require alternating legend sides for visual variety. | User‑customizable dashboards that let end‑users flip legend orientation with a single action.
// AI Prompts: Generate C# code with Aspose.Cells that detects a chart legend's current Position and moves it to the opposite corner, then saves the workbook. | Provide an Aspose.Cells snippet that toggles a legend between left/right or top/bottom, handling default cases gracefully. | Create a reusable C# method that accepts a Chart object and flips its legend position using LegendPositionType.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart, reads the legend's current Position, switches it to the opposite side (left↔right, top↔bottom, otherwise Corner), and saves the file as LegendPositionSwitched.xlsx using Aspose.Cells.
class LegendPositionSwitcher
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the legend
        Legend legend = chart.Legend;

        // Read current legend position
        LegendPositionType currentPos = legend.Position;

        // Determine opposite corner position
        LegendPositionType newPos;
        switch (currentPos)
        {
            case LegendPositionType.Left:
                newPos = LegendPositionType.Right;
                break;
            case LegendPositionType.Right:
                newPos = LegendPositionType.Left;
                break;
            case LegendPositionType.Top:
                newPos = LegendPositionType.Bottom;
                break;
            case LegendPositionType.Bottom:
                newPos = LegendPositionType.Top;
                break;
            default:
                // For Corner, NotDocked or any other value, default to Corner
                newPos = LegendPositionType.Corner;
                break;
        }

        // Apply the new position
        legend.Position = newPos;

        // Save the workbook with the updated legend position
        workbook.Save("LegendPositionSwitched.xlsx");
    }
}
