using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

class ConvertRangeToListObjectAndChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data (including headers)
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Value");
        cells["A2"].PutValue("A");
        cells["B2"].PutValue(10);
        cells["A3"].PutValue("B");
        cells["B3"].PutValue(20);
        cells["A4"].PutValue("C");
        cells["B4"].PutValue(30);

        // Convert the data range into a ListObject (table) so that charts can auto‑update
        // Add the ListObject using the string‑based overload (startCell, endCell, hasHeaders)
        int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.DisplayName = "DataTable";   // optional: give the table a friendly name

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the chart's data range to the table's data range.
        // Using the table's DataRange ensures the chart updates automatically when the table changes.
        chart.SetChartDataRange(table.DataRange.Address, true); // true = plot series by rows (vertical)

        // Save the workbook
        workbook.Save("ChartWithTable.xlsx", SaveFormat.Xlsx);
    }
}