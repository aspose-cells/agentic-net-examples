using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class InsertPieChartExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new worksheet that will contain the pie chart
        Worksheet chartSheet = workbook.Worksheets[workbook.Worksheets.Add()];

        // Populate sample data in the range A1:A5 (values for the pie chart)
        chartSheet.Cells["A1"].PutValue(10);
        chartSheet.Cells["A2"].PutValue(20);
        chartSheet.Cells["A3"].PutValue(30);
        chartSheet.Cells["A4"].PutValue(25);
        chartSheet.Cells["A5"].PutValue(15);

        // Add a pie chart to the worksheet.
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = chartSheet.Charts.Add(ChartType.Pie, 2, 1, 20, 10);
        Chart pieChart = chartSheet.Charts[chartIndex];

        // Set the data source for the pie chart (values only)
        pieChart.NSeries.Add("A1:A5", true);

        // Optional: set a title for the chart
        pieChart.Title.Text = "Sample Pie Chart";

        // Save the workbook to a file
        workbook.Save("PieChartOnNewSheet.xlsx", SaveFormat.Xlsx);
    }
}