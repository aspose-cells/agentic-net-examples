using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

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
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);          // Values
        chart.NSeries.CategoryData = "A2:A3";      // Categories

        // Define a red‑component threshold for the series color
        const int redThreshold = 150;

        // Iterate through each series in the chart
        for (int i = 0; i < chart.NSeries.Count; i++)
        {
            Series series = chart.NSeries[i];

            // Ensure the series has a foreground color; assign one if not set
            if (series.Area.ForegroundColor.IsEmpty || series.Area.ForegroundColor.ToArgb() == 0)
            {
                // Example: set a reddish color for demonstration purposes
                series.Area.ForegroundColor = Color.FromArgb(200, 50, 50);
            }

            Color seriesColor = series.Area.ForegroundColor;

            // If the series color's red component exceeds the threshold,
            // remove the fill of the corresponding legend entry text
            if (seriesColor.R > redThreshold)
            {
                LegendEntry legendEntry = series.LegendEntry;
                legendEntry.IsTextNoFill = true;   // No fill for legend text
            }
        }

        // Save the workbook to a file
        workbook.Save("ChartLegendConditional.xlsx");
    }
}