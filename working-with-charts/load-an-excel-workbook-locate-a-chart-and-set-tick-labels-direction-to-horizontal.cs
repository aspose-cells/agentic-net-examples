// Title: Aspose.Cells for .NET – Set Chart Axis Tick Labels to Horizontal (C#)
// Description: Load an Excel workbook, locate a chart, and change the tick‑label orientation of both the category (X) and value (Y) axes to horizontal using Aspose.Cells for .NET, then save the file.
// Keywords: Aspose.Cells chart axis label direction | C# set tick labels horizontal | Excel chart formatting Aspose | CategoryAxis TickLabels Horizontal | ValueAxis TickLabels Horizontal | Aspose.Cells axis text orientation | .NET Excel chart customization
// Common Searches: Aspose.Cells change chart axis label orientation | C# set chart tick labels horizontal Aspose | How to make Excel chart axis labels horizontal with Aspose.Cells | Set category axis tick labels to horizontal in .NET | Modify value axis label direction Aspose.Cells
// Developer Intent: Adjust the orientation of a chart’s axis tick labels to horizontal in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Enhance readability of axis labels in automated financial reports. | Apply a consistent horizontal label style across all charts before publishing dashboards. | Meet corporate branding guidelines that require non‑rotated chart labels.
// AI Prompts: Generate C# code that iterates through every chart in a workbook and sets both X‑ and Y‑axis tick labels to horizontal with Aspose.Cells. | Show how to change chart axis tick labels to vertical for a selected chart using Aspose.Cells for .NET. | Explain how to reset chart axis tick label direction to the default setting in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load an Excel workbook, locate a chart, and change the tick‑label orientation of both the category (X) and value (Y) axes to horizontal using Aspose.Cells for .NET, then save the file.
class SetTickLabelsDirection
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (sheet.Charts.Count > 0)
        {
            // Get the first chart in the worksheet
            Chart chart = sheet.Charts[0];

            // Set the tick labels direction of the category (X) axis to horizontal
            chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Horizontal;

            // Optionally, also set the value (Y) axis tick labels direction to horizontal
            chart.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Horizontal;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
