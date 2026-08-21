// Title: C# – Conditional Legend Entry Fill in Aspose.Cells Column Chart Based on Series Color
// Description: Creates a workbook, adds sample data, builds a column chart with two series, assigns explicit colors, defines a threshold color, and sets the legend entry's IsTextNoFill property to true when the series color matches the threshold. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells C# chart legend | conditional legend fill | IsTextNoFill property | series color threshold | column chart example | Excel legend formatting | Aspose.Cells sample code
// Common Searches: Aspose.Cells set legend entry no fill C# | conditional legend formatting based on series color | how to hide legend fill when series is red Aspose.Cells | C# example for chart legend conditional rule | Aspose.Cells chart legend IsTextNoFill usage
// Developer Intent: Apply a rule that removes the legend entry fill when a series color equals a predefined threshold.
// Use Cases: Highlight critical series in financial dashboards by hiding their legend fill when they exceed a risk threshold. | Generate automated reports where negative performance values are colored red and their legend entries appear without fill for quick visual scanning. | Create dynamic Excel charts that adapt legend styling based on real‑time color thresholds across multiple data series.
// AI Prompts: Generate C# code using Aspose.Cells that sets IsTextNoFill on legend entries when the series foreground color matches a given Color. | Show how to loop through chart series in Aspose.Cells and apply a conditional legend fill rule based on a threshold color, then save the workbook. | Explain how to extend the example to support several threshold colors and different legend styling options in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, builds a column chart with two series, assigns explicit colors, defines a threshold color, and sets the legend entry's IsTextNoFill property to true when the series color matches the threshold. The workbook is then saved as an Excel file.
class LegendEntryConditionalFill
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(60);
        sheet.Cells["C3"].PutValue(130);
        sheet.Cells["C4"].PutValue(90);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series to the chart
        chart.NSeries.Add("B2:B4", true); // Series 1
        chart.NSeries.Add("C2:C4", true); // Series 2
        chart.NSeries.CategoryData = "A2:A4";

        // Define a color threshold (example: Red)
        Color thresholdColor = Color.Red;

        // Assign explicit colors to series for demonstration
        chart.NSeries[0].Area.ForegroundColor = Color.Red;      // Matches threshold
        chart.NSeries[1].Area.ForegroundColor = Color.Blue;    // Does not match

        // Iterate through each series and apply conditional rule
        for (int i = 0; i < chart.NSeries.Count; i++)
        {
            Series series = chart.NSeries[i];
            // Check if the series foreground color matches the threshold
            if (series.Area.ForegroundColor.ToArgb() == thresholdColor.ToArgb())
            {
                // Set legend entry text fill to none
                series.LegendEntry.IsTextNoFill = true;
            }
            else
            {
                // Ensure normal fill for other entries
                series.LegendEntry.IsTextNoFill = false;
            }
        }

        // Save the workbook
        workbook.Save("LegendEntryConditionalFill.xlsx");
    }
}
