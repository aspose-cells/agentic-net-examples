using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AssignSeriesToSecondaryAxis
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
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Secondary");
        sheet.Cells["C2"].PutValue(100);
        sheet.Cells["C3"].PutValue(200);
        sheet.Cells["C4"].PutValue(300);

        // Add a line chart
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series: primary and secondary
        chart.NSeries.Add("B2:B4", true); // primary series
        chart.NSeries.Add("C2:C4", true); // secondary series
        chart.NSeries.CategoryData = "A2:A4";

        // Assign the second series to the secondary vertical axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Optional: set a title for the secondary axis
        chart.SecondValueAxis.Title.Text = "Secondary Axis";

        // Save the workbook
        workbook.Save("LineSeriesSecondaryAxis.xlsx");
    }
}