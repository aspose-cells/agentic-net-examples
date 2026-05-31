using Aspose.Cells;
using Aspose.Cells.Charts;

class SecondaryYAxisDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        sheet.Cells["B1"].PutValue("Primary");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(200);

        sheet.Cells["C1"].PutValue("Secondary");
        sheet.Cells["C2"].PutValue(5000);
        sheet.Cells["C3"].PutValue(3000);
        sheet.Cells["C4"].PutValue(1000);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series: first uses primary Y‑axis, second will use secondary Y‑axis
        chart.NSeries.Add("B2:B4", true); // primary series
        chart.NSeries.Add("C2:C4", true); // secondary series
        chart.NSeries.CategoryData = "A2:A4";

        // Link the second series to the secondary Y‑axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Configure the secondary Y‑axis (value axis)
        Axis secondaryValueAxis = chart.SecondValueAxis;
        secondaryValueAxis.Title.Text = "Secondary Values";
        secondaryValueAxis.MinValue = 0;
        secondaryValueAxis.MaxValue = 6000;
        secondaryValueAxis.MajorUnit = 1000;

        // Save the workbook
        workbook.Save("SecondaryYAxisDemo.xlsx");
    }
}