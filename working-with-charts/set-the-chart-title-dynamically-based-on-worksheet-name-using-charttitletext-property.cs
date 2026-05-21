using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChartTitle
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Optionally rename the worksheet to demonstrate dynamic title
            worksheet.Name = "SalesData";

            // Add sample data for the chart
            worksheet.Cells["A1"].PutValue("Month");
            worksheet.Cells["A2"].PutValue("Jan");
            worksheet.Cells["A3"].PutValue("Feb");
            worksheet.Cells["A4"].PutValue("Mar");
            worksheet.Cells["B1"].PutValue("Revenue");
            worksheet.Cells["B2"].PutValue(15000);
            worksheet.Cells["B3"].PutValue(18000);
            worksheet.Cells["B4"].PutValue(21000);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Dynamically set the chart title to the worksheet name
            chart.Title.Text = worksheet.Name;
            chart.Title.IsVisible = true;

            // Save the workbook
            workbook.Save("DynamicChartTitle.xlsx");
        }
    }
}