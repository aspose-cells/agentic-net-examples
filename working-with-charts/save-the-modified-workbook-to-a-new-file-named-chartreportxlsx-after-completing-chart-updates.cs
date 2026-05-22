using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Value1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a second data series
        worksheet.Cells["C1"].PutValue("Value2");
        worksheet.Cells["C2"].PutValue(15);
        worksheet.Cells["C3"].PutValue(25);
        worksheet.Cells["C4"].PutValue(35);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the first series and category data
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Add the second series to the chart
        chart.NSeries.Add("C2:C4", true);

        // Save the modified workbook to a new file named ChartReport.xlsx
        workbook.Save("ChartReport.xlsx", SaveFormat.Xlsx);
    }
}