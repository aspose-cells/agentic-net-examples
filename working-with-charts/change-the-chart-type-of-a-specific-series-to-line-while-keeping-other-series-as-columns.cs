using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ChangeSeriesChartType
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate category data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["A5"].PutValue("Q4");

        // Populate first series data (will be changed to line)
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(40);

        // Populate second series data (remains column)
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);
        sheet.Cells["C5"].PutValue(45);

        // Add a column chart (default type for all series)
        int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 15);
        Chart chart = sheet.Charts[chartIdx];

        // Add both series to the chart
        chart.NSeries.Add("B2:B5", true); // first series
        chart.NSeries.Add("C2:C5", true); // second series
        chart.NSeries.CategoryData = "A2:A5";

        // Change the first series to a line chart type while keeping the second as column
        chart.NSeries[0].Type = ChartType.Line;

        // Optional: customize the line appearance of the changed series
        chart.NSeries[0].Border.Color = Color.Red;

        // Save the workbook
        workbook.Save("SeriesTypeChanged.xlsx");
    }
}