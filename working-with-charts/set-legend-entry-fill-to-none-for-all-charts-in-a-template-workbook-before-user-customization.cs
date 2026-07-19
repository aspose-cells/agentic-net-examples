// Title: Aspose.Cells .NET – Remove Fill from All Chart Legend Entries in a Workbook
// Description: Load a template workbook, loop through every worksheet and chart, ensure the legend is visible, and set each LegendEntry's IsTextNoFill property to true (including entries from LegendEntries collection). Save the workbook with transparent legend backgrounds ready for further styling or export.
// Keywords: Aspose.Cells legend no fill | C# chart legend transparent background | remove legend entry fill Aspose.Cells | iterate charts workbook Aspose.Cells | set IsTextNoFill true | Aspose.Cells .NET chart styling | global chart legend formatting | US Aspose.Cells examples | UK chart legend customization | Canada workbook legend transparency
// Common Searches: Aspose.Cells set legend entry IsTextNoFill for all charts | C# remove legend background from every chart in a workbook | how to make chart legends transparent using Aspose.Cells | iterate worksheets and charts to clear legend fill .NET | template workbook legend styling Aspose.Cells
// Developer Intent: Apply a transparent fill to every legend entry across all charts in a workbook before any additional customization.
// Use Cases: Standardize legend appearance in corporate report templates to match brand guidelines. | Prepare charts for PDF export where legend shading interferes with readability. | Automate uniform legend styling for large workbooks generated from data pipelines.
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, iterates all worksheets and charts, ensures legends are shown, and sets LegendEntry.IsTextNoFill = true for each series and any LegendEntries collection, then saves the file. | Show an example of clearing legend entry fill for every chart type in a template workbook using Aspose.Cells for .NET. | Explain how to handle charts that expose legend entries via chart.Legend.LegendEntries when removing legend fill with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Load a template workbook, loop through every worksheet and chart, ensure the legend is visible, and set each LegendEntry's IsTextNoFill property to true (including entries from LegendEntries collection). Save the workbook with transparent legend backgrounds ready for further styling or export.
class SetLegendEntryNoFill
{
    static void Main()
    {
        // Load the template workbook
        Workbook workbook = new Workbook("Template.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts in the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Ensure the legend is visible (optional, but safe)
                chart.ShowLegend = true;

                // Set no fill for legend entries associated with each series
                foreach (Series series in chart.NSeries)
                {
                    LegendEntry legendEntry = series.LegendEntry;
                    legendEntry.IsTextNoFill = true;
                }

                // Some chart types expose legend entries via the LegendEntries collection.
                // Apply the same setting to any entries found there.
                LegendEntryCollection legendEntries = chart.Legend.LegendEntries;
                if (legendEntries != null)
                {
                    foreach (LegendEntry entry in legendEntries)
                    {
                        entry.IsTextNoFill = true;
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("Output.xlsx");
    }
}
