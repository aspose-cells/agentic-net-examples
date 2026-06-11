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

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and configure them
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;               // show the values
        series.DataLabels.Position = LabelPositionType.InsideEnd; // position of the labels

        // Freeze the rows that contain the data (first 4 rows) so they stay visible while scrolling
        sheet.FreezePanes(5, 0, 4, 0); // freeze top 4 rows

        // Save the workbook
        workbook.Save("ChartWithDataLabelsAndFreeze.xlsx", SaveFormat.Xlsx);
    }
}