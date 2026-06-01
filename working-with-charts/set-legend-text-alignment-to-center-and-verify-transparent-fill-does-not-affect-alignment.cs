using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

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
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(80);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);          // Values
        chart.NSeries.CategoryData = "A2:A3";      // Categories

        // Access the chart legend
        Legend legend = chart.Legend;

        // Center align the legend text horizontally and vertically
        legend.TextHorizontalAlignment = TextAlignmentType.Center;
        legend.TextVerticalAlignment = TextAlignmentType.Center;

        // Set the legend background to transparent.
        // Transparent background should not affect the text alignment.
        legend.BackgroundMode = BackgroundMode.Transparent;

        // Save the workbook
        workbook.Save("LegendAlignmentTransparent.xlsx");
    }
}