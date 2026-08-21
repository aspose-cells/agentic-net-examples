// Title: Add and style leader lines for a chart series in an existing XLSX using Aspose.Cells for .NET
// Description: Loads an XLSX workbook, accesses the first worksheet and its first chart, enables leader lines for the first series, disables automatic formatting, and applies a dotted blue line (1.5 pt) before saving the file.
// Keywords: Aspose.Cells chart leader lines | C# modify Excel chart series | enable leader lines Aspose | customize chart leader line style | .NET Excel chart formatting | leader line color weight
// Common Searches: Aspose.Cells add leader lines to chart series | C# set leader line style in Excel chart | how to customize chart leader lines with Aspose | change leader line color and weight in .NET Excel | enable leader lines for first chart in workbook
// Developer Intent: Programmatically turn on leader lines for the first series of the first chart in an existing workbook and define their visual properties.
// Use Cases: Enhance a sales trend chart by highlighting data points with blue dotted leader lines. | Prepare presentation‑ready workbooks where the primary chart needs clear leader lines for readability. | Automate consistent chart styling across multiple reports by applying the same leader‑line settings.
// AI Prompts: Generate C# code that uses Aspose.Cells to enable leader lines for every series in all charts of a workbook, setting a solid red line of 2 pt weight. | Create a reusable method that accepts a file path and leader‑line parameters (style, weight, color) and applies them to the first series of the first chart.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsLeaderLinesExample
{
    // Loads an XLSX workbook, accesses the first worksheet and its first chart, enables leader lines for the first series, disables automatic formatting, and applies a dotted blue line (1.5 pt) before saving the file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the first chart in the worksheet
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Get the first series
            Series series = chart.NSeries[0];

            // Enable leader lines for the series
            series.HasLeaderLines = true;

            // Customize leader line properties
            series.LeaderLines.IsAuto = false;                     // Disable automatic formatting
            series.LeaderLines.Style = LineType.Dot;               // Set line style to dotted
            series.LeaderLines.WeightPt = 1.5;                     // Set line weight (points)
            series.LeaderLines.Color = Color.Blue;                // Set line color

            // Save the modified workbook (replace with desired output path)
            workbook.Save("output.xlsx");

            Console.WriteLine("Leader line settings updated and workbook saved.");
        }
    }
}
