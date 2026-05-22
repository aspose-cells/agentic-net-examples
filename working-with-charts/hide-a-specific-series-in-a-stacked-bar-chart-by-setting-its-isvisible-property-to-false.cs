using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class HideSeriesInStackedBarChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for categories and two series
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

        // Add a stacked bar chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 15, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Add the two series to the chart
        chart.NSeries.Add("B2:B4", true); // Series1
        chart.NSeries.Add("C2:C4", true); // Series2
        chart.NSeries.CategoryData = "A2:A4";

        // Hide the second series (Series2) by marking it as filtered
        chart.NSeries[1].IsFiltered = true;

        // Save the workbook to a file
        workbook.Save("StackedBar_HideSeries.xlsx");
    }
}