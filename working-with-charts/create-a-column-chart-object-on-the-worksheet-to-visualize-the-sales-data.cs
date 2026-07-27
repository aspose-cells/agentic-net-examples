using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample sales data
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["B1"].PutValue("Sales");

        string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
        int[] sales = { 1200, 1500, 1100, 1800, 1600 };

        for (int i = 0; i < months.Length; i++)
        {
            sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A
            sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B
        }

        // Add a column chart to the worksheet (using the Add(ChartType, int, int, int, int) overload)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (A1:B6) and plot by column
        chart.SetChartDataRange("A1:B6", true);

        // Set a title for the chart
        chart.Title.Text = "Monthly Sales";

        // Save the workbook with the chart
        workbook.Save("SalesChart.xlsx", SaveFormat.Xlsx);
    }
}