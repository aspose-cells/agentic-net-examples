// Title: Read and toggle a chart legend’s position to the opposite corner using Aspose.Cells for .NET (C#)
// AI Prompts: Retrieve the current LegendPosition of the first chart in a worksheet and set it to the opposite corner with Aspose.Cells in C#. | Switch a chart legend between top/bottom or left/right positions programmatically and save the workbook using Aspose.Cells for .NET. | Change a chart legend from a docked position to NotDocked or Corner based on its existing setting and write the updated Excel file with Aspose.Cells.
// Common Searches: how to change chart legend position to opposite corner using Aspose.Cells C# | Aspose.Cells read legend position and move it programmatically | toggle Excel chart legend top bottom left right with Aspose.Cells .NET | C# code to switch chart legend from docked to not docked in Aspose.Cells | save workbook after updating chart legend position Aspose.Cells
// Tags: Aspose.Cells chart legend repositioning | C# toggle legend position Aspose.Cells | Excel chart legend corner switch .NET | Aspose.Cells modify legend docking | save workbook after legend update Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook, reads the legend position of the first chart, switches it to the opposite side or corner, updates the legend, and saves the modified file as a new workbook.
class LegendPositionSwitcher
{
    static void Main()
    {
        // Load an existing workbook that contains a chart
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one chart
        if (sheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Access the first chart
        Chart chart = sheet.Charts[0];

        // Access the legend of the chart
        Legend legend = chart.Legend;

        // Read the current legend position
        LegendPositionType currentPos = legend.Position;
        Console.WriteLine($"Current legend position: {currentPos}");

        // Determine the opposite corner position
        LegendPositionType newPos;
        switch (currentPos)
        {
            case LegendPositionType.Top:
                newPos = LegendPositionType.Bottom;
                break;
            case LegendPositionType.Bottom:
                newPos = LegendPositionType.Top;
                break;
            case LegendPositionType.Left:
                newPos = LegendPositionType.Right;
                break;
            case LegendPositionType.Right:
                newPos = LegendPositionType.Left;
                break;
            case LegendPositionType.Corner:
                // If positioned in a corner, switch to the opposite corner (Top-Left <-> Bottom-Right, etc.)
                // For simplicity, toggle between Corner and NotDocked
                newPos = LegendPositionType.NotDocked;
                break;
            default:
                // For NotDocked or any other value, set to a default corner
                newPos = LegendPositionType.Corner;
                break;
        }

        // Apply the new position
        legend.Position = newPos;
        Console.WriteLine($"Legend position changed to: {newPos}");

        // Save the workbook with the updated legend position
        workbook.Save("output.xlsx");
    }
}
