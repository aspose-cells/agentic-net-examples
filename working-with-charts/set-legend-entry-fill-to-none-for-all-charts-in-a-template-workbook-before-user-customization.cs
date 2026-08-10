// Title: Aspose.Cells C# – Remove Legend Entry Fill for All Charts in a Workbook Template
// Description: Loads a template workbook, loops through every worksheet and chart, ensures the legend is visible, sets each legend entry’s text fill to none, and saves the updated file. Ideal for preparing chart templates with transparent legend text.
// Keywords: Aspose.Cells legend no fill | C# chart legend entry transparent | remove legend background Aspose.Cells | set legend entry IsTextNoFill | iterate charts workbook Aspose.Cells | .NET Excel chart formatting | template workbook legend styling
// Common Searches: Aspose.Cells set legend entry no fill for all charts | C# remove legend background color in Excel charts | how to make legend text transparent using Aspose.Cells | loop through worksheets and charts to clear legend fill | Aspose.Cells chart legend formatting example
// Developer Intent: The developer needs to clear any fill color from legend entry text across every chart in a template workbook before applying further customizations.
// Use Cases: Create a clean chart template where legend text has no background, allowing downstream styling. | Generate automated reports that require legends to appear without fill on multiple sheets. | Standardize chart appearance in a workbook by enforcing no‑fill legend entries before distribution.
// AI Prompts: Generate C# code with Aspose.Cells that sets IsTextNoFill = true for every legend entry in all charts of a workbook. | Show an Aspose.Cells example that iterates through worksheets and charts to remove legend fill while keeping the legend visible. | Explain how to extend the code to also remove legend borders for each chart in the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a template workbook, loops through every worksheet and chart, ensures the legend is visible, sets each legend entry’s text fill to none, and saves the updated file. Ideal for preparing chart templates with transparent legend text.
class Program
{
    static void Main()
    {
        // Load the template workbook
        Workbook workbook = new Workbook("Template.xlsx");

        // Loop through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through all charts in the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Ensure the legend is visible (optional)
                chart.ShowLegend = true;

                // Set no fill for legend entry text of each series
                foreach (Series series in chart.NSeries)
                {
                    series.LegendEntry.IsTextNoFill = true;
                }

                // Also set no fill for any additional legend entries
                if (chart.Legend.LegendEntries != null)
                {
                    foreach (LegendEntry entry in chart.Legend.LegendEntries)
                    {
                        entry.IsTextNoFill = true;
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("Result.xlsx");
    }
}
