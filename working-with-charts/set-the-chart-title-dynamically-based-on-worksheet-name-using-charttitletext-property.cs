using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetChartTitleDynamic
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and give it a meaningful name
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = "SalesData";

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Dynamically set the chart title to the worksheet's name
        chart.Title.Text = worksheet.Name;

        // Save the workbook to a file
        workbook.Save("DynamicChartTitle.xlsx");
    }
}