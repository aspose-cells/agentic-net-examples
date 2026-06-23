using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetDataLabelCurrencyFormat
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
        worksheet.Cells["B2"].PutValue(1000);
        worksheet.Cells["B3"].PutValue(2000);
        worksheet.Cells["B4"].PutValue(3000);

        // Add a bar chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Bar, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;

        // Set the data label number format to currency with two decimal places
        series.DataLabels.NumberFormat = "$#,##0.00";

        // Save the workbook to a file
        workbook.Save("BarChartDataLabelsCurrency.xlsx");
    }
}