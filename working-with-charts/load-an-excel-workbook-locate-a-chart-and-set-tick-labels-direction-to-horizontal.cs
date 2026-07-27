// Title: Set Chart Axis Tick‑Label Direction to Horizontal with Aspose.Cells for .NET
// Description: Loads an Excel workbook, retrieves the first worksheet and its first chart, sets the Category (X) axis tick‑label orientation to horizontal via Aspose.Cells, and saves the modified file.
// Keywords: Aspose.Cells C# chart axis label orientation | horizontal tick labels Aspose.Cells | set CategoryAxis TickLabels direction | Excel chart label rotation .NET | programmatically change axis label direction
// Common Searches: Aspose.Cells set X axis labels horizontal | C# change chart tick label direction Aspose.Cells | how to rotate chart axis labels in Excel using Aspose.Cells | make chart axis text horizontal .NET | Aspose.Cells chart label orientation example
// Developer Intent: Change a chart axis tick‑label orientation to horizontal in an Excel workbook using Aspose.Cells.
// Use Cases: Enhance readability of dense category labels in automatically generated reports. | Apply a consistent horizontal label style across multiple workbooks for branding purposes. | Adjust axis orientation dynamically when column widths are modified during workbook creation.
// AI Prompts: Write C# code with Aspose.Cells that sets both Category and Value axis tick‑label directions to vertical for every chart in a workbook. | Show how to iterate through all worksheets and charts in an Excel file and apply a horizontal tick‑label orientation to the X‑axis. | Explain how to conditionally choose the tick‑label direction based on label length when using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook, retrieves the first worksheet and its first chart, sets the Category (X) axis tick‑label orientation to horizontal via Aspose.Cells, and saves the modified file.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one chart
        if (worksheet.Charts.Count > 0)
        {
            // Get the first chart in the worksheet
            Chart chart = worksheet.Charts[0];

            // Set the tick‑label direction of the category (X) axis to horizontal
            chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Horizontal;

            // If you need to set the value (Y) axis instead, uncomment the line below
            // chart.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Horizontal;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
