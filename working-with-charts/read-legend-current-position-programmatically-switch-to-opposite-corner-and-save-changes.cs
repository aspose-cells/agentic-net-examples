// Title: C# – Read and Toggle Chart Legend Position (Opposite Corner) with Aspose.Cells
// Description: Loads a workbook, accesses the first chart, reads its Legend.Position, switches left↔right or top↔bottom (corner stays unchanged, other values default to Corner), updates the legend, and saves the file.
// Keywords: Aspose.Cells C# chart legend position | legend.Position Aspose.Cells | toggle chart legend side | move Excel chart legend programmatically | set legend to corner Aspose.Cells | C# read legend placement | Aspose.Cells modify chart legend
// Common Searches: Aspose.Cells get legend position C# | Aspose.Cells set legend to right | switch chart legend left to right Aspose | C# change Excel chart legend location | Aspose.Cells legend corner toggle
// Developer Intent: Read a chart’s legend placement, change it to the opposite side or corner, and save the workbook.
// Use Cases: Automatically improve readability of generated Excel reports by standardizing legend locations. | Apply consistent legend placement across all charts before distributing a workbook. | Enable a UI button that flips legend positions without manual Excel editing.
// AI Prompts: Generate C# code using Aspose.Cells to read a chart legend’s position and move it to the opposite side, handling all LegendPositionType values. | Explain how to detect the current Legend.Position and apply a fallback to Corner when the legend is not docked. | Show how to iterate through every chart in a workbook, toggle each legend to the opposite corner, and save the updated file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a workbook, accesses the first chart, reads its Legend.Position, switches left↔right or top↔bottom (corner stays unchanged, other values default to Corner), updates the legend, and saves the file.
class LegendPositionSwitcher
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one chart in the worksheet
        if (sheet.Charts.Count == 0)
        {
            Console.WriteLine("No charts found in the worksheet.");
            return;
        }

        // Get the first chart
        Chart chart = sheet.Charts[0];

        // Access the legend of the chart
        Legend legend = chart.Legend;

        // Read the current legend position
        LegendPositionType currentPos = legend.Position;
        Console.WriteLine($"Current Legend Position: {currentPos}");

        // Determine the opposite corner position
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
            case LegendPositionType.Corner:
                // If already at Corner, toggle to opposite corner (still Corner)
                newPos = LegendPositionType.Corner;
                break;
            default:
                // For NotDocked or any other value, set to Corner as a fallback
                newPos = LegendPositionType.Corner;
                break;
        }

        // Apply the new position
        legend.Position = newPos;
        Console.WriteLine($"New Legend Position set to: {newPos}");

        // Save the workbook with changes
        workbook.Save("output.xlsx");
        Console.WriteLine("Workbook saved as output.xlsx");
    }
}
