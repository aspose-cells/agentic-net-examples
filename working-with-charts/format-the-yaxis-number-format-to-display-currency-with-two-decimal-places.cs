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
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(1234.56);
        worksheet.Cells["B3"].PutValue(2345.67);
        worksheet.Cells["B4"].PutValue(3456.78);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Format the Y‑axis (ValueAxis) tick labels as currency with two decimal places
        chart.ValueAxis.TickLabels.NumberFormat = "$#,##0.00";

        // Save the workbook to a file
        workbook.Save("YAxisCurrencyFormat.xlsx");
    }
}