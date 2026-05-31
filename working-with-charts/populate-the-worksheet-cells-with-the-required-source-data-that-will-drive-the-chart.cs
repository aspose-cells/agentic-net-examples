using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class PopulateChartData
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["C1"].PutValue("Series2");

        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(20);

        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["C3"].PutValue(40);

        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(50);
        sheet.Cells["C4"].PutValue(60);

        // Add a column chart that will use the populated data
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data range (vertical series)
        chart.SetChartDataRange("A1:C4", true);

        // Save the workbook
        workbook.Save("ChartData.xlsx", SaveFormat.Xlsx);
    }
}