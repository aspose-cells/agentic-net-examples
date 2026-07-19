// Title: C# – Modify Leader Lines of the First Chart Series in an Existing XLSX Workbook with Aspose.Cells
// Description: Load an XLSX file, locate the first worksheet's first chart, enable leader lines for its initial data series, and customize style, weight, and color using Aspose.Cells before saving the result.
// Keywords: Aspose.Cells C# chart leader lines | modify Excel chart series style | set leader line color Aspose | customize chart leader line weight | enable leader lines in XLSX with .NET
// Common Searches: Aspose.Cells enable leader lines for chart series | C# change leader line style in Excel chart | how to set dotted leader lines in Aspose.Cells | modify chart series leader line color .NET | Aspose.Cells leader line weight example
// Developer Intent: Activate leader lines for the first data series of the first chart in a loaded workbook and apply custom formatting (style, thickness, color) via Aspose.Cells.
// Use Cases: Improve label clarity in pie or doughnut charts by adding blue dotted leader lines. | Apply a consistent leader‑line appearance across multiple reports generated from a template workbook. | Automate chart styling in batch processing where every workbook must follow corporate visual standards.
// AI Prompts: Generate C# code with Aspose.Cells that sets dash‑style red leader lines for all series in every chart of a workbook. | Create a method that iterates through all worksheets and charts, enabling leader lines and applying a 2‑point weight uniformly. | Provide robust error handling for scenarios where a worksheet lacks charts or a chart lacks series while modifying leader line properties.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

// Load an XLSX file, locate the first worksheet's first chart, enable leader lines for its initial data series, and customize style, weight, and color using Aspose.Cells before saving the result.
class ModifyLeaderLines
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Check that the worksheet contains at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Retrieve the first chart in the collection
            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one data series
            if (chart.NSeries.Count > 0)
            {
                // Get the first series of the chart
                Series series = chart.NSeries[0];

                // Enable leader lines for the series
                series.HasLeaderLines = true;

                // Customize the leader lines appearance
                series.LeaderLines.IsAuto = false;               // Disable automatic formatting
                series.LeaderLines.Style = LineType.Dot;         // Set line style to dotted
                series.LeaderLines.WeightPt = 1.5;               // Set line weight (points)
                series.LeaderLines.Color = Color.Blue;           // Set line color
            }
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
