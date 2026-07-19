// Title: Set GapWidth for a column series in Aspose.Cells .NET to adjust bar spacing
// Description: Creates a workbook, adds sample data, inserts a 2‑D column chart, sets the first series' GapWidth to 150 % (controlling column spacing), and saves the file as AdjustedGapWidth.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | .NET chart gap width | column chart bar spacing | Chart.NSeries GapWidth | adjust column spacing Aspose.Cells | set column chart gap width | gap width percentage
// Common Searches: Aspose.Cells change column chart gap width | C# set GapWidth for chart series | increase spacing between bars Aspose.Cells | how to adjust column chart bar spacing .NET | GapWidth property example Aspose.Cells
// Developer Intent: Modify the GapWidth property of a chart series to control the distance between columns.
// Use Cases: Design column charts with custom bar spacing for better readability. | Generate reports where chart density must match corporate visual guidelines. | Create multiple charts with varying gap widths to compare visual impact.
// AI Prompts: Show how to set GapWidth to 200 % for a column series in Aspose.Cells C#. | Provide code that changes GapWidth for each series in a stacked column chart using Aspose.Cells. | Explain the valid GapWidth range and its visual effect on column charts in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AdjustGapWidthDemo
{
    // Creates a workbook, adds sample data, inserts a 2‑D column chart, sets the first series' GapWidth to 150 % (controlling column spacing), and saves the file as AdjustedGapWidth.xlsx using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a 2‑D column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Adjust the gap width of the first (and only) series.
            // This controls the spacing between column clusters.
            // Value is a percentage of the column width (0‑500). 150 = 150%.
            chart.NSeries[0].GapWidth = 150;

            // Save the workbook with the modified chart
            workbook.Save("AdjustedGapWidth.xlsx");
        }
    }
}
