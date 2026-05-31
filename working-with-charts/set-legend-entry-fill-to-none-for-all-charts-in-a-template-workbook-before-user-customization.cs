using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

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
                // Ensure the chart has a legend and series
                if (chart.Legend != null && chart.NSeries != null)
                {
                    // Iterate through each series to access its legend entry
                    foreach (Series series in chart.NSeries)
                    {
                        LegendEntry legendEntry = series.LegendEntry;
                        if (legendEntry != null)
                        {
                            // Set the legend entry text to have no fill
                            legendEntry.IsTextNoFill = true;
                        }
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("Template_Modified.xlsx");
    }
}