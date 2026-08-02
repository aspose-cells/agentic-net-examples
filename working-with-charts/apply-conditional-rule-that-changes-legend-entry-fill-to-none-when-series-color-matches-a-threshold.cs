// Title: Hide chart legend entry fill when series color meets a threshold using Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds a column chart, sets the first series color to a threshold (red), checks the series color, and if it matches the threshold sets the legend entry's IsTextNoFill property to true, producing a legend without fill before saving the file.
// Keywords: Aspose.Cells | C# | chart legend | IsTextNoFill | conditional formatting | threshold color | column chart | Excel automation | legend entry fill | series color | Aspose.Cells API
// Common Searches: Aspose.Cells hide legend fill based on series color | C# set LegendEntry.IsTextNoFill | Conditional legend formatting Aspose.Cells | Chart series color threshold Aspose.Cells | Remove legend entry fill in Excel using Aspose.Cells
// Developer Intent: Apply a conditional rule that removes the fill of a chart legend entry when the associated series color equals a predefined threshold.
// Use Cases: Generate a sales chart where red bars indicate values above a limit and the legend entry for the red series is shown without fill for visual clarity. | Process multiple chart series programmatically and automatically suppress legend fills for any series that match out‑of‑range colors. | Create dashboards that dynamically hide legend fills when colors are driven by data‑dependent thresholds.
// AI Prompts: Write C# code with Aspose.Cells that loops through all chart series and sets LegendEntry.IsTextNoFill to true when the series ForegroundColor equals Color.Red. | Show how to read a threshold color from a worksheet cell and apply it to conditionally remove legend entry fill in an Aspose.Cells chart. | Explain how to revert a legend entry's fill after using IsTextNoFill, including code to restore default styling.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example creates a workbook, adds a column chart, sets the first series color to a threshold (red), checks the series color, and if it matches the threshold sets the legend entry's IsTextNoFill property to true, producing a legend without fill before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(80);
        sheet.Cells["B3"].PutValue(120);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Define the threshold color
        Color thresholdColor = Color.Red;

        // Apply the threshold color to the first series
        chart.NSeries[0].Area.ForegroundColor = thresholdColor;

        // Access the legend entry associated with the first series
        LegendEntry legendEntry = chart.NSeries[0].LegendEntry;

        // If the series color matches the threshold, remove fill from the legend text
        if (chart.NSeries[0].Area.ForegroundColor.ToArgb() == thresholdColor.ToArgb())
        {
            legendEntry.IsTextNoFill = true; // No fill for legend entry text
        }

        // Save the workbook to a file
        workbook.Save("LegendEntryConditionalFill.xlsx");
    }
}
