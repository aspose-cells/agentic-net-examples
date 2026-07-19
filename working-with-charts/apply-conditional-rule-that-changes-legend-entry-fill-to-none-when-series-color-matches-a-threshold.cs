// Title: Conditional Legend Entry Fill Based on Series Color in Aspose.Cells (C#)
// Description: Creates a workbook, adds a column chart with a colored series, defines a red‑component threshold, checks the series foreground color, and sets the legend entry's IsTextNoFill property when the threshold is exceeded, then saves the file.
// Keywords: Aspose.Cells | C# | chart legend | conditional formatting | legend entry fill | IsTextNoFill | series color threshold | Excel chart API | column chart | dynamic legend styling
// Common Searches: Aspose.Cells set legend entry no fill C# | conditional legend formatting based on series color | how to hide legend fill when color is red Aspose.Cells | IsTextNoFill property example | apply color threshold to chart legend Aspose
// Developer Intent: Remove fill from a chart legend entry when the series color exceeds a defined red component threshold.
// Use Cases: Suppress legend fill for warning‑level series colored bright red in financial dashboards. | Automatically adjust legend styling for high‑risk data points during report generation. | Create clean visualizations where intense colors indicate alerts and legend text should remain unfilled.
// AI Prompts: Generate C# code with Aspose.Cells that loops through all chart series and applies IsTextNoFill to legend entries whose red component is greater than 200. | Show an example that changes the legend entry background instead of text fill when a series' green component exceeds a threshold. | Explain how to implement multiple conditional legend styles for different series colors in a single Aspose.Cells chart.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart with a colored series, defines a red‑component threshold, checks the series foreground color, and sets the legend entry's IsTextNoFill property when the threshold is exceeded, then saves the file.
class LegendEntryConditionalFill
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
        sheet.Cells["A4"].PutValue("Q3");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Assign a specific color to the series (example: bright red)
        chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(255, 10, 10); // Very red

        // Define a color threshold (e.g., Red component > 200)
        const int redThreshold = 200;

        // Retrieve the series color
        Color seriesColor = chart.NSeries[0].Area.ForegroundColor;

        // If the series color meets the threshold, remove fill from the legend entry text
        if (seriesColor.R > redThreshold)
        {
            // Access the legend entry associated with the first series
            LegendEntry legendEntry = chart.NSeries[0].LegendEntry;

            // Set IsTextNoFill to true so the legend text has no fill
            legendEntry.IsTextNoFill = true;
        }

        // Save the workbook
        workbook.Save("LegendEntryConditionalFill.xlsx");
    }
}
