using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ExportChartSeriesToJson
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Fill sample data: categories in column A and values in column B
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add a chart and bind the data series to column B
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);               // Y‑values
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";          // X‑axis categories

        // Export the chart series data to JSON by iterating the source column (B)
        int startRow = 1;                         // Zero‑based index for row 2 (first data row)
        int endRow = sheet.Cells.MaxDataRow;      // Last row that contains data
        int valueColumn = 1;                      // Column B (zero‑based)

        StringBuilder jsonBuilder = new StringBuilder();
        jsonBuilder.Append("[");

        for (int row = startRow; row <= endRow; row++)
        {
            string category = sheet.Cells[row, 0].StringValue;   // Column A
            double value = sheet.Cells[row, valueColumn].DoubleValue;

            jsonBuilder.Append("{");
            jsonBuilder.AppendFormat("\"Category\":\"{0}\",\"Value\":{1}", category, value);
            jsonBuilder.Append("}");

            if (row < endRow)
                jsonBuilder.Append(",");
        }

        jsonBuilder.Append("]");

        string jsonResult = jsonBuilder.ToString();
        Console.WriteLine(jsonResult);

        // Save the workbook (optional)
        workbook.Save("ChartSeries.xlsx");
    }
}