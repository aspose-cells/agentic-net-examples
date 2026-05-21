using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AddSecondaryYAxisDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for categories and two series
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Series 1");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["B4"].PutValue(300);

        sheet.Cells["C1"].PutValue("Series 2");
        sheet.Cells["C2"].PutValue(5000);
        sheet.Cells["C3"].PutValue(3000);
        sheet.Cells["C4"].PutValue(1000);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series to the chart and set the category (X‑axis) data
        chart.NSeries.Add("B2:B4", true);   // Series 1
        chart.NSeries.Add("C2:C4", true);   // Series 2
        chart.NSeries.CategoryData = "A2:A4";

        // Render the second series on the secondary Y‑axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Optional: customize the secondary Y‑axis appearance
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.Title.Text = "Secondary Axis";
        secondaryAxis.MinValue = 0;
        secondaryAxis.MaxValue = 6000;
        secondaryAxis.MajorUnit = 1000;
        secondaryAxis.IsVisible = true;

        // Save the workbook with the chart
        workbook.Save("ChartWithSecondaryYAxis.xlsx");
    }
}