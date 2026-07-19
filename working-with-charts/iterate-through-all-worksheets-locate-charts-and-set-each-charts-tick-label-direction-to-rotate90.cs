// Title: Set Every Chart Axis Tick Label to Rotate90 Across All Worksheets with Aspose.Cells for .NET
// Description: Creates or loads a Workbook, loops through each worksheet and every chart, and applies ChartTextDirectionType.Rotate90 to both the Category (X) and Value (Y) axes tick labels, then saves the file.
// Keywords: Aspose.Cells rotate chart axis labels | ChartTextDirectionType Rotate90 | C# set tick label direction | iterate worksheets charts Aspose | vertical axis labels .NET | programmatic chart formatting | Aspose.Cells axis label orientation
// Common Searches: rotate chart axis labels 90 degrees Aspose.Cells | C# loop through worksheets and change chart tick label direction | set vertical tick labels for all charts in a workbook | Aspose.Cells change axis label orientation programmatically
// Developer Intent: Apply a 90‑degree rotation to the tick labels of both category and value axes for every chart in all worksheets of a workbook.
// Use Cases: Standardize vertical axis labels in multi‑sheet financial reports before exporting to PDF. | Improve readability of dense charts in printed workbooks by rotating tick labels. | Automate chart styling for dashboards that require consistent label orientation across sheets.
// AI Prompts: Generate a C# method that receives a Workbook and sets ChartTextDirectionType.Rotate90 on all chart axes, handling charts without a value axis. | Explain how to check for the existence of a CategoryAxis before modifying its TickLabels in Aspose.Cells. | Provide code to rotate axis tick labels to 90° for every chart in a workbook and then export the workbook as a PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates or loads a Workbook, loops through each worksheet and every chart, and applies ChartTextDirectionType.Rotate90 to both the Category (X) and Value (Y) axes tick labels, then saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts on the current worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Set tick label direction for the category (X) axis to Rotate90
                chart.CategoryAxis.TickLabels.DirectionType = ChartTextDirectionType.Rotate90;

                // Set tick label direction for the value (Y) axis to Rotate90
                chart.ValueAxis.TickLabels.DirectionType = ChartTextDirectionType.Rotate90;
            }
        }

        // Save the modified workbook
        workbook.Save("Output.xlsx");
    }
}
