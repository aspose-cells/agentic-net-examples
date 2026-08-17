// Title: C# – Set Legend Entry Fill to Transparent for Negative Series in Aspose.Cells Column Chart
// Description: Creates a workbook with positive and mixed (including negative) data series, adds a column chart, and uses LegendEntry.BackgroundMode to make the legend of the series that contains negative values transparent while keeping other legends opaque. An optional red font highlights the negative series before saving the file.
// Keywords: Aspose.Cells legend entry transparent | BackgroundMode.Transparent C# | chart legend customization Aspose.Cells | negative values legend Aspose | column chart legend fill Aspose.Cells
// Common Searches: Aspose.Cells set legend entry transparent for negative series | how to change legend background mode in Aspose.Cells chart | C# make chart legend entry transparent based on data | Aspose.Cells legend font color red for negative values
// Developer Intent: Apply a transparent background to the legend entry of any chart series that includes negative data points, leaving other legend entries unchanged.
// Use Cases: Visually distinguish a series with negative values by making its legend entry transparent and coloring its text red. | Generate Excel reports where legends reflect data polarity without altering the chart colors. | Automate chart styling in bulk workbooks, applying transparent legends only to series that contain negatives.
// AI Prompts: Generate C# code using Aspose.Cells that scans all chart series and sets LegendEntry.BackgroundMode to Transparent for series with at least one negative value, also changing the legend font color to red. | Explain the difference between BackgroundMode.Opaque and BackgroundMode.Transparent for LegendEntry in Aspose.Cells and show how to toggle them based on series data. | Provide a step‑by‑step tutorial for customizing legend fills in Aspose.Cells column charts, including handling of negative data points.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook with positive and mixed (including negative) data series, adds a column chart, and uses LegendEntry.BackgroundMode to make the legend of the series that contains negative values transparent while keeping other legends opaque. An optional red font highlights the negative series before saving the file.
class SetLegendEntryFillForNegativeValues
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data:
        // Series 1 – all positive values
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["B1"].PutValue("Positive");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(70);
        sheet.Cells["B4"].PutValue(60);

        // Series 2 – contains negative values
        sheet.Cells["C1"].PutValue("Negative");
        sheet.Cells["C2"].PutValue(-30);
        sheet.Cells["C3"].PutValue(20);
        sheet.Cells["C4"].PutValue(-10);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add the two series to the chart
        // First series (positive values)
        chart.NSeries.Add("B2:B4", true);
        // Second series (mixed values, includes negatives)
        chart.NSeries.Add("C2:C4", true);
        // Set category (X‑axis) labels
        chart.NSeries.CategoryData = "A2:A4";

        // Access legend entries for each series
        LegendEntry positiveLegend = chart.NSeries[0].LegendEntry;
        LegendEntry negativeLegend = chart.NSeries[1].LegendEntry;

        // Keep the positive series legend opaque (default)
        positiveLegend.BackgroundMode = BackgroundMode.Opaque;

        // Set the legend entry for the series that has negative values to transparent
        negativeLegend.BackgroundMode = BackgroundMode.Transparent;

        // Optionally, customize the font color to highlight the change
        negativeLegend.Font.Color = Color.Red;

        // Save the workbook
        workbook.Save("LegendEntryNegativeTransparent.xlsx");
    }
}
