using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDashboard
{
    class HideLegendDemo
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet (dashboard sheet)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart that occupies a large area of the sheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 2, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend to maximize the display area for the chart
            chart.ShowLegend = false;

            // Save the workbook
            workbook.Save("DashboardChart_NoLegend.xlsx");
        }
    }
}