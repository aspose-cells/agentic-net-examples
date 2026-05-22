using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["A2"].PutValue("Cat1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Cat2");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("Cat3");
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Assign the data source to the chart using SetChartDataRange
        // The second parameter 'true' indicates plotting by column (vertical)
        chart.SetChartDataRange("A1:B4", true);

        // Save the workbook to a file
        workbook.Save("ColumnChartWithDataRange.xlsx");
    }
}