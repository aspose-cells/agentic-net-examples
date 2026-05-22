using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class HideLegendCharts
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the charts
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Generate a series of charts (e.g., three charts) and hide their legends
        for (int i = 0; i < 3; i++)
        {
            // Determine chart position to avoid overlap
            int upperRow = 5 + i * 15;
            int lowerRow = upperRow + 10;

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, upperRow, 0, lowerRow, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend to reduce visual noise in the slideshow
            chart.ShowLegend = false;
        }

        // Save the workbook with the charts
        workbook.Save("ChartsWithoutLegend.xlsx");
    }
}