using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetChartDataRangeExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells with sample data
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Series1");
        worksheet.Cells["A2"].PutValue("Cat1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Cat2");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("Cat3");
        worksheet.Cells["B4"].PutValue(30);

        // Add a chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the chart's data source range (vertical series)
        chart.SetChartDataRange("A1:B4", true);

        // Save the workbook
        workbook.Save("ChartWithDataRange.xlsx");
    }
}