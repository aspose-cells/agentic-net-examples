using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class StackedBarCustomColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a stacked bar chart
        // Categories
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");

        // Series 1
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Series 2
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Series 3
        sheet.Cells["D1"].PutValue("Series3");
        sheet.Cells["D2"].PutValue(12);
        sheet.Cells["D3"].PutValue(22);
        sheet.Cells["D4"].PutValue(32);

        // Add a stacked bar chart
        int chartIdx = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 15);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the series (B2:D4) and categories (A2:A4)
        chart.NSeries.Add("B2:D4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Assign custom colors to each series using the Area.ForegroundColor property
        // (Series.Style is not directly exposed; Area.ForegroundColor effectively sets the series fill color)
        chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(79, 129, 189);   // Blueish
        chart.NSeries[1].Area.ForegroundColor = Color.FromArgb(192, 80, 77);    // Reddish
        chart.NSeries[2].Area.ForegroundColor = Color.FromArgb(155, 187, 89);   // Greenish

        // Optionally set border colors for better visual distinction
        chart.NSeries[0].Border.Color = Color.Black;
        chart.NSeries[1].Border.Color = Color.Black;
        chart.NSeries[2].Border.Color = Color.Black;

        // Save the workbook
        workbook.Save("StackedBarCustomColors.xlsx");
    }
}