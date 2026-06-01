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

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Quarter");
        sheet.Cells["B1"].PutValue("Revenue");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["B2"].PutValue(15000);
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B3"].PutValue(20000);
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["B4"].PutValue(18000);
        sheet.Cells["A5"].PutValue("Q4");
        sheet.Cells["B5"].PutValue(22000);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Set the chart title text and apply bold formatting
        chart.Title.Text = "Quarterly Revenue";
        chart.Title.IsVisible = true;
        chart.Title.Font.IsBold = true;

        // Save the workbook to a file
        workbook.Save("QuarterlyRevenueChart.xlsx");
    }
}