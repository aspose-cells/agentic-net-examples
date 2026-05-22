using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class LockChartAspectRatio
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);          // Set values
        chart.NSeries.CategoryData = "A2:A4";      // Set categories

        // Lock the chart's aspect ratio to prevent distortion when resizing the plot area
        chart.ChartObject.IsAspectRatioLocked = true;

        // Save the workbook with the locked chart
        workbook.Save("ChartAspectRatioLocked.xlsx");
    }
}