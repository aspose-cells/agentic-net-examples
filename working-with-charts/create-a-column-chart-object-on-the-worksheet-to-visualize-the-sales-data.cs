using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample sales data
            // Header
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");

            // Data rows
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            int[] sales = { 12000, 15000, 13000, 17000, 16000, 18000 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B
            }

            // Add a column chart to the worksheet (topRow, leftColumn, bottomRow, rightColumn)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart (including headers)
            chart.SetChartDataRange("A1:B7", true); // true => plot by column

            // Optional: set chart title and enable legend
            chart.Title.Text = "Monthly Sales";
            chart.ShowLegend = true;

            // Save the workbook
            workbook.Save("SalesColumnChart.xlsx", SaveFormat.Xlsx);
        }
    }
}