// Title: Read and Switch Chart Legend Position with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a column chart, reads the current LegendPositionType, flips it to the opposite corner (left↔right, top↔bottom), updates the legend, and saves the file as LegendPositionSwitched.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | chart legend position | LegendPositionType | switch legend corner | programmatic legend move | Excel chart manipulation | save workbook
// Common Searches: Aspose.Cells get chart legend position | change legend corner C# Aspose | move Excel chart legend to opposite side programmatically | how to toggle legend position with Aspose.Cells | read and set legend position in .NET
// Developer Intent: Programmatically detect a chart legend’s location, move it to the opposite corner, and persist the change.
// Use Cases: Automatically adjust legend placement for reports that switch between portrait and landscape layouts. | Create multilingual templates that flip legends for left‑to‑right versus right‑to‑left reading directions. | Batch‑process workbooks to enforce a consistent opposite‑corner legend style across all charts.
// AI Prompts: Generate C# code with Aspose.Cells that reads a chart legend’s current position, switches it to the opposite corner, and saves the workbook. | Show how to handle LegendPositionType.Corner and NotDocked when toggling legend locations in an Excel file using Aspose.Cells. | Explain step‑by‑step how to programmatically move a chart legend from top to bottom (or left to right) in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart, reads the current LegendPositionType, flips it to the opposite corner (left↔right, top↔bottom), updates the legend, and saves the file as LegendPositionSwitched.xlsx using Aspose.Cells.
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
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the legend
        Legend legend = chart.Legend;

        // Read the current legend position
        LegendPositionType currentPos = legend.Position;
        Console.WriteLine($"Current legend position: {currentPos}");

        // Determine the opposite corner position
        LegendPositionType newPos = currentPos switch
        {
            LegendPositionType.Left => LegendPositionType.Right,
            LegendPositionType.Right => LegendPositionType.Left,
            LegendPositionType.Top => LegendPositionType.Bottom,
            LegendPositionType.Bottom => LegendPositionType.Top,
            // For Corner and NotDocked keep the same (no opposite defined)
            _ => currentPos
        };

        // Apply the new position
        legend.Position = newPos;
        Console.WriteLine($"New legend position set to: {newPos}");

        // Save the workbook with the updated legend position
        workbook.Save("LegendPositionSwitched.xlsx");
    }
}
