using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class StackedColumnChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        // Column A: Categories
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["A5"].PutValue("Q4");

        // Column B: Series 1 values
        sheet.Cells["B1"].PutValue("Series 1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(40);

        // Column C: Series 2 values
        sheet.Cells["C1"].PutValue("Series 2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);
        sheet.Cells["C5"].PutValue(45);

        // Column D: Series 3 values
        sheet.Cells["D1"].PutValue("Series 3");
        sheet.Cells["D2"].PutValue(12);
        sheet.Cells["D3"].PutValue(22);
        sheet.Cells["D4"].PutValue(32);
        sheet.Cells["D5"].PutValue(42);

        // Add a stacked column chart
        int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 6, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add the three series (B:D) and set category data (A)
        chart.NSeries.Add("B2:D5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Customize each series color individually
        chart.NSeries[0].Area.ForegroundColor = Color.Red;      // Series 1
        chart.NSeries[1].Area.ForegroundColor = Color.Green;    // Series 2
        chart.NSeries[2].Area.ForegroundColor = Color.Blue;     // Series 3

        // Save the workbook
        workbook.Save("StackedColumnChart.xlsx");
    }
}