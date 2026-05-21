using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that will be used for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Convert the range A1:B4 into a ListObject (table) so that chart updates automatically
        int tableIndex = sheet.ListObjects.Add("A1", "B4", true); // hasHeaders = true
        ListObject table = sheet.ListObjects[tableIndex];
        table.DisplayName = "DataTable";

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart's data range to the table's data range (including headers)
        chart.SetChartDataRange(table.DataRange.Address, true); // true = plot by column

        // Save the workbook
        workbook.Save("ListObjectChart.xlsx", SaveFormat.Xlsx);
    }
}