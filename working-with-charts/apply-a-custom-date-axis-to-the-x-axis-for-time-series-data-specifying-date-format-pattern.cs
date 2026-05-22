using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with date (X) and numeric (Y) data
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
        sheet.Cells["B4"].PutValue(300);

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data source: values and category (date) data
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure the X (category) axis to use a time scale
        chart.CategoryAxis.CategoryType = CategoryType.TimeScale;

        // Apply a custom date format to the X‑axis values
        chart.NSeries[0].XValuesFormatCode = "yyyy-MM-dd";

        // Save the workbook
        workbook.Save("CustomDateAxisChart.xlsx");
    }
}